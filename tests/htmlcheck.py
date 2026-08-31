"""Execute report.html in a real browser and assert it actually rendered.

The Markdown renderers run in PowerShell, but the HTML report is built by
JavaScript from an embedded JSON payload. A PowerShell array that collapses to
a scalar produces valid JSON and a silently broken page - the docs look fine
while half the report is blank. Only executing the page catches that.
"""

import sys
from playwright.sync_api import sync_playwright

# id -> (needle it must contain, JS guard that says whether there IS data).
# A section with nothing to show is correctly empty; only a section that has
# data and still renders blank is a failure.
REQUIRED = {
    "kpis":              ("card",     "true"),
    "ca-table":          ("<tr",      "(D.ConditionalAccess.Policies||[]).length > 0"),
    "ca-gaps":           ("badge",    "(D.CaGaps||[]).length > 0"),
    "roles":             ("bar-row",  "(D.Roles||[]).length > 0"),
    "licenses":          ("bar-row",  "(D.Licenses||[]).length > 0"),
    "group-tiles":       ("card",     "true"),
    "auth-table":        ("<tr",      "true"),
    "settings-table":    ("<tr",      "true"),
    "intune-compliance": ("<tr",      "!!(D.Intune && D.Intune.Available)"),
}


def check(path):
    errors, failures = [], []
    with sync_playwright() as p:
        b = p.chromium.launch()
        pg = b.new_page()
        pg.on("pageerror", lambda e: errors.append(str(e)))
        pg.on("console", lambda m: errors.append(m.text) if m.type == "error" else None)
        pg.goto("file://" + path)
        pg.wait_for_timeout(500)
        checked = 0
        for el_id, (needle, guard) in REQUIRED.items():
            if not pg.query_selector("#" + el_id):
                failures.append("#%s missing from the document" % el_id)
                continue
            if not pg.evaluate("() => " + guard):
                continue  # nothing to render here, correctly blank
            checked += 1
            html = pg.eval_on_selector("#" + el_id, "e => e.innerHTML")
            if needle not in html:
                failures.append("#%s has data but rendered empty (expected %r)" % (el_id, needle))
        # Arrays that must survive the PowerShell -> JSON round trip as arrays.
        bad = pg.evaluate("""() => {
            const out = [];
            (D.ConditionalAccess.Policies || []).forEach((p, i) => {
              ['IncludeUsers','ExcludeUsers','IncludeGroups','ExcludeGroups',
               'IncludeApps','ExcludeApps'].forEach(f => {
                 if (p[f] != null && !Array.isArray(p[f])) out.push('policy ' + i + '.' + f);
              });
              ['Include','Exclude'].forEach(f => {
                 const v = (p.Locations || {})[f];
                 if (v != null && !Array.isArray(v)) out.push('policy ' + i + '.Locations.' + f);
              });
            });
            ((D.Intune || {}).CompliancePolicies || []).forEach((p, i) => {
              if (p.Assignments != null && !Array.isArray(p.Assignments)) out.push('compliance ' + i + '.Assignments');
            });
            ((D.Intune || {}).ConfigurationProfiles || []).forEach((p, i) => {
              if (p.Assignments != null && !Array.isArray(p.Assignments)) out.push('profile ' + i + '.Assignments');
            });
            (D.AuthMethods || []).forEach((m, i) => {
              if (m.Targets != null && !Array.isArray(m.Targets)) out.push('authmethod ' + i + '.Targets');
            });
            return out;
        }""")
        for f in bad:
            failures.append("%s collapsed from an array to a scalar" % f)
        b.close()
    return errors, failures, checked


def main():
    rc = 0
    for path in sys.argv[1:]:
        errors, failures, checked = check(path)
        label = path.split("/")[-2] + "/" + path.split("/")[-1]
        if errors:
            print("FAIL %s: %d JavaScript error(s)" % (label, len(errors)))
            for e in errors[:5]:
                print("       %s" % e)
            rc = 1
        elif failures:
            print("FAIL %s: page loaded but sections did not render" % label)
            for f in failures:
                print("       %s" % f)
            rc = 1
        else:
            print("PASS %s: no JS errors, %d populated section(s) rendered, arrays intact"
                  % (label, checked))
    return rc


if __name__ == "__main__":
    sys.exit(main())
