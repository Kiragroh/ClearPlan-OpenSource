"""Static (AST) and runtime enforcement of the read-only RayStation boundary."""
import ast
import builtins
import os
import re
import unittest

import _paths
import fake_raystation as fr
import raystation_adapter
import raystation_patient_loader

RS = _paths.RAYSTATION
API_FILES = ["raystation_adapter.py", "raystation_patient_loader.py", "run_live_review.py",
             "run_retrospective_mining.py"]
CORE = os.path.join(RS, "clearplan_review_core")
WRITE_LIKE = re.compile(r"^(Save|Create|Delete|Update|Set|Add|Remove|Import|Export|Compute|Recompute|Run|Apply|Copy|"
                        r"Edit|Execute|Unload|Close|Merge|Approve|Unapprove|Change|Modify|Store|Commit|Write|Move|"
                        r"Rename|Clear|Reset|Optimize|Simplify|Expand|Contract|Margin|Algebra)[A-Z0-9_]")
ALLOWED = {
    "raystation_adapter.py": raystation_adapter.ALLOWED_RAYSTATION_METHODS,
    "raystation_patient_loader.py": raystation_patient_loader.ALLOWED_RAYSTATION_METHODS,
    "run_live_review.py": frozenset(),
    "run_retrospective_mining.py": frozenset(),
}
ROI_ALGEBRA_OR_DOSE = ("ComputeDose", "CreateAlgebraGeometry", "CreateMarginGeometry", "SetAlgebraExpression",
                       "ComputeRoiGeometry", "UpdateDependency", "ComputeDoseOnAdditionalSets", "Save")


def tree(name):
    with open(os.path.join(RS, name), encoding="utf-8") as handle:
        source = handle.read()
    return source, ast.parse(source)


class StaticBoundary(unittest.TestCase):
    def test_only_allowlisted_pascal_case_calls(self):
        for name in API_FILES:
            _, module = tree(name)
            for node in ast.walk(module):
                if isinstance(node, ast.Call) and isinstance(node.func, ast.Attribute):
                    attr = node.func.attr
                    if attr[:1].isupper():
                        self.assertIn(attr, ALLOWED[name], "%s calls %s" % (name, attr))

    def test_no_write_like_names_anywhere(self):
        for name in API_FILES:
            source, module = tree(name)
            for node in ast.walk(module):
                text = None
                if isinstance(node, ast.Attribute):
                    text = node.attr
                elif isinstance(node, ast.Constant) and isinstance(node.value, str) and \
                        re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", node.value):
                    text = node.value  # identifier-like strings (getattr names, keyword dicts), not messages
                elif isinstance(node, ast.Name) and not hasattr(builtins, node.id):
                    text = node.id
                if text is not None:
                    self.assertIsNone(WRITE_LIKE.match(text), "%s uses write-like name %r" % (name, text))
            for forbidden in ROI_ALGEBRA_OR_DOSE:
                self.assertNotIn(forbidden + "(", source, name)

    def test_no_attribute_assignment_setattr_or_delete(self):
        for name in API_FILES:
            _, module = tree(name)
            for node in ast.walk(module):
                targets = []
                if isinstance(node, ast.Assign):
                    targets = node.targets
                elif isinstance(node, (ast.AugAssign, ast.AnnAssign)):
                    targets = [node.target]
                elif isinstance(node, ast.Delete):
                    targets = node.targets
                for target in targets:
                    for sub in ast.walk(target):
                        if isinstance(sub, ast.Attribute):
                            self.assertFalse(sub.attr[:1].isupper(),
                                             "%s assigns RayStation-like attribute %s" % (name, sub.attr))
                if isinstance(node, ast.Call) and isinstance(node.func, ast.Name):
                    self.assertNotIn(node.func.id, ("setattr", "delattr", "exec", "eval"), name)

    def test_patient_database_is_isolated_in_the_loader(self):
        for name in API_FILES:
            source, _ = tree(name)
            if name != "raystation_patient_loader.py":
                self.assertNotIn("PatientDB", source, name)
                self.assertNotIn("LoadPatient", source, name)
        source, _ = tree("raystation_patient_loader.py")
        self.assertIn("AllowPatientUpgrade=False", source)
        self.assertIn("VERIFY LOCALLY", source)

    def test_core_never_touches_a_tps_api(self):
        core_classes = {"DictReader", "StringIO"}  # csv.DictReader, io.StringIO (standard library)
        for file_name in os.listdir(CORE):
            if file_name.endswith(".py"):
                with open(os.path.join(CORE, file_name), encoding="utf-8") as handle:
                    core_classes.update(n.name for n in ast.walk(ast.parse(handle.read())) if isinstance(n, ast.ClassDef))
        for file_name in os.listdir(CORE):
            if not file_name.endswith(".py"):
                continue
            with open(os.path.join(CORE, file_name), encoding="utf-8") as handle:
                module = ast.parse(handle.read())
            for node in ast.walk(module):
                if isinstance(node, (ast.Import, ast.ImportFrom)):
                    names = [a.name for a in node.names] + ([node.module] if getattr(node, "module", None) else [])
                    for imported in names:
                        self.assertFalse(imported.split(".")[0] in ("raystation", "connect", "clr", "System"),
                                         "%s imports %s" % (file_name, imported))
                if isinstance(node, ast.Call) and isinstance(node.func, ast.Attribute):
                    attr = node.func.attr
                    self.assertFalse(attr[:1].isupper() and attr not in core_classes, "%s calls %s" % (file_name, attr))


class RuntimeBoundary(unittest.TestCase):
    def test_live_read_uses_only_allowlisted_methods(self):
        del fr.CALLS[:]
        _, get_current = fr.build_phantom()
        plan = raystation_adapter.read_current_plan(get_current)
        from clearplan_review_core import engine, goal_table
        table = goal_table.load(os.path.join(_paths.RULES, "example_goal_table.json"))
        engine.evaluate(plan, table, now_utc=fr.REVIEW_TIME)
        methods = set(c for c in fr.CALLS if not c.startswith("get_current:"))
        self.assertTrue(methods)
        self.assertTrue(methods <= raystation_adapter.ALLOWED_RAYSTATION_METHODS, methods)
        self.assertEqual({"get_current:Patient", "get_current:Case", "get_current:Plan", "get_current:BeamSet",
                          "get_current:Examination"}, set(c for c in fr.CALLS if c.startswith("get_current:")))

    def test_loader_never_allows_patient_upgrade(self):
        objects, _ = fr.build_phantom()
        database, get_current = fr.build_database([objects["Patient"]])
        del fr.CALLS[:]
        raystation_patient_loader.open_plan_readonly(get_current, "RS-PHANTOM-001", "Case 1", "1AA_PCA_CP", "")
        self.assertEqual([("RS-PHANTOM-001", False)], database.load_calls)
        self.assertEqual({"QueryPatientInfo", "LoadPatient"},
                         set(c for c in fr.CALLS if not c.startswith("get_current:")))


if __name__ == "__main__":
    unittest.main()
