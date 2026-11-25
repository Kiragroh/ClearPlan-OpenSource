"""Audit actual native output, without changing results or replacing missing values."""
import csv, json, hashlib, pathlib, collections, sys
HERE=pathlib.Path(__file__).resolve().parents[1]
stage=json.loads((HERE/'staging.json').read_text(encoding='utf-8'))
root=pathlib.Path(stage['protected_results'])
tables={name:list(csv.DictReader((root/'review_log'/f'{name}.csv').open(encoding='utf-8-sig',newline='')))
        for name in ['review','goal_result','plan_parameter','check_finding','ruleset_revision']}
reviews=tables['review']
out={'native_review_count':len(reviews),'runs':[],'comparisons':[],'assertions':[]}
def children(name,r): return [x for x in tables[name] if x['review_id']==r['review_id']]
def semantic(name,r):
    return [{k:v for k,v in x.items() if k!='review_id'} for x in children(name,r)]
for r in reviews:
    revisions=children('ruleset_revision',r)
    rev=next((x['revision_number'] for x in revisions if x['config_key']=='constraints-refdb'),None)
    goals=children('goal_result',r)
    configured=[x for x in goals if x['template_code'].startswith(('901','902'))]
    reports=[p for p in (root/'reports').rglob('*.pdf') if hashlib.sha256(p.read_bytes()).hexdigest()==r['report_sha256']] if r['report_written']=='1' else []
    out['runs'].append(dict(review_id=r['review_id'],mode=r['mode'],fractions=int(r['fractions']) if r['fractions'] else None,
        revision=rev,status=r['run_status'],goals=len(goals),configured_goals=len(configured),
        configured_states=dict(collections.Counter(x['status'] for x in configured)),
        checks=len(children('check_finding',r)),parameters=len(children('plan_parameter',r)),
        ruleset_hash=r['ruleset_sha256'],goal_table_hash=r['goal_table_sha256'],
        report_hash_matches=len(reports),report_written=r['report_written'],
        duration_seconds=None,missing_structure_null=all(x['achieved_value']=='' and x['status']=='not_evaluated' for x in configured if x['template_structure']=='PHANTOM_MISSING')))
    from datetime import datetime
    out['runs'][-1]['duration_seconds']=(datetime.fromisoformat(r['finished_utc'].replace('Z','+00:00'))-datetime.fromisoformat(r['started_utc'].replace('Z','+00:00'))).total_seconds()
for fx in ['5','40']:
    rs=[r for r in reviews if r['fractions']==fx]
    def pick(mode,rev):
        matches=[r for r in rs if r['mode']==mode and any(x['config_key']=='constraints-refdb' and x['revision_number']==rev for x in children('ruleset_revision',r))]
        return matches[-1] if matches else None
    a,b,c=pick('headless','1'),pick('headless','3'),pick('retrospective','1')
    if a and b:
        ag,bg=children('goal_result',a),children('goal_result',b)
        changed=[]
        for x,y in zip(ag,bg):
            diff={k:[x[k],y[k]] for k in x if k!='review_id' and x[k]!=y[k]}
            if diff:changed.append(dict(code=x['template_code'],changes=diff))
        out['comparisons'].append(dict(fractions=int(fx),kind='rule-change',
            measurement_values_identical=[x['achieved_value'] for x in ag]==[x['achieved_value'] for x in bg],
            ruleset_changed=a['ruleset_sha256']!=b['ruleset_sha256'],goal_table_changed=a['goal_table_sha256']!=b['goal_table_sha256'],
            changed_goals=changed,parameters_equal=semantic('plan_parameter',a)==semantic('plan_parameter',b),checks_equal=semantic('check_finding',a)==semantic('check_finding',b)))
    if a and c:
        out['comparisons'].append(dict(fractions=int(fx),kind='retrospective-parity',
            goals_equal=semantic('goal_result',a)==semantic('goal_result',c),parameters_equal=semantic('plan_parameter',a)==semantic('plan_parameter',c),
            checks_equal=semantic('check_finding',a)==semantic('check_finding',c),ruleset_equal=a['ruleset_sha256']==c['ruleset_sha256'],
            goal_table_equal=a['goal_table_sha256']==c['goal_table_sha256'],pseudonym_equal=a['plan_key']==c['plan_key'],
            counts={name:len(children(name,a)) for name in ['goal_result','plan_parameter','check_finding']}))
    live=pick('live','1')
    if a and live:
        out['comparisons'].append(dict(fractions=int(fx),kind='live-close-parity',
            goals_equal=semantic('goal_result',a)==semantic('goal_result',live),parameters_equal=semantic('plan_parameter',a)==semantic('plan_parameter',live),
            checks_equal=semantic('check_finding',a)==semantic('check_finding',live),ruleset_equal=a['ruleset_sha256']==live['ruleset_sha256'],
            no_report=live['report_written']=='0'))
out['all_rows_have_parent']=all(x['review_id'] in {r['review_id'] for r in reviews} for name in tables if name!='review' for x in tables[name])
out['counts']={name:len(rows) for name,rows in tables.items()}
(HERE/'native_verification.json').write_text(json.dumps(out,indent=2,ensure_ascii=False),encoding='utf-8')
print(json.dumps(out,indent=2,ensure_ascii=False))
