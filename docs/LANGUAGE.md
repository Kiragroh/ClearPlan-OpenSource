# Review display language

The shared review workspace has a `de` / `en` selector. Fixed controls and
presentation labels update in place; DVH series identity, values and visibility
are retained. German is the default. The preference is stored per Windows account
under local application data in `ClearPlan/review-language.txt`; optional
`CLEARPLAN_LANGUAGE=de` or `en` selects the startup preference.

PDF and HTML capture the display-language code when the detached report document
is created. Rendering temporarily uses that captured language and restores the
current GUI language afterward. The stable title remains **Plan Quality Report**.

This is presentation localization, not translation of clinical records. Patient,
plan, field and structure identifiers, numeric measurements, stored status codes,
and unmatched free diagnostic text remain unchanged. Unknown source wording can
therefore remain in its original language. Local rules and clinical terminology
still need local review; a translated label does not change an evaluation.

Regression tests exercise the real shared selector, nested language scopes,
view construction, plot identity and GUI/HTML/PDF language round trips using
artificial records. They do not establish complete language coverage or clinical
acceptance. Run `tools/test-portable-review.ps1` for the portable software suite.
