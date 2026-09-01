"""Leak scan for Export-EntraTenantDocs.ps1 -Anonymize.

Harvests every identifying string from the source snapshot AND every history
snapshot the run will read, then greps every rendered output for them. Any hit
is a leak. Also asserts the things that are supposed to survive still do -
an anonymizer that redacts everything would pass a leak scan and be useless.
"""

import json
import os
import re
import sys

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Literals the tool deliberately keeps: Microsoft vocabulary, not identity.
KEEP_PRINCIPALS = {"All", "None", "GuestsOrExternalUsers"}
KEEP_APPS = {"All cloud apps", "None", "Office 365", "Microsoft Admin Portals",
             "All", "Office365", "MicrosoftAdminPortals"}
KEEP_LOCATIONS = {"All", "AllTrusted"}
KEEP_TARGETS = {"all_users", "all_devices"}
KEEP_ASSIGNMENTS = {"All devices", "All users"}

# Short or generic strings would collide with ordinary report text.
MIN_LEN = 4


def token_in(value, text):
    """Whole-token, case-insensitive: the scrub replaces case-insensitively,
    so the scanner must look the same way - and a short value like 'main'
    inside 'domain' must not count."""
    pat = r'(?<![A-Za-z0-9])' + re.escape(value) + r'(?![A-Za-z0-9])'
    return re.search(pat, text, re.IGNORECASE) is not None


def add(bag, value, why, allow=frozenset()):
    if value is None:
        return
    s = str(value).strip()
    if not s or s in allow or len(s) < MIN_LEN:
        return
    bag.setdefault(s, set()).add(why)


def harvest(d, bag):
    """Every value that MUST NOT appear in anonymized output."""
    add(bag, d.get("TenantId"), "TenantId")
    org = d.get("Organization") or {}
    add(bag, org.get("DisplayName"), "Organization.DisplayName")
    for dom in org.get("Domains") or []:
        add(bag, dom.get("Name"), "domain")

    ca = d.get("ConditionalAccess") or {}
    for p in ca.get("Policies") or []:
        for f in ("IncludeUsers", "ExcludeUsers"):
            for v in p.get(f) or []:
                add(bag, v, "CA " + f, KEEP_PRINCIPALS)
        for f in ("IncludeGroups", "ExcludeGroups"):
            for v in p.get(f) or []:
                add(bag, v, "CA " + f, KEEP_PRINCIPALS)
        for f in ("IncludeApps", "ExcludeApps"):
            for v in p.get(f) or []:
                add(bag, v, "CA " + f, KEEP_APPS)
        for f in ("Include", "Exclude"):
            for v in (p.get("Locations") or {}).get(f) or []:
                add(bag, v, "CA location", KEEP_LOCATIONS)
    for l in ca.get("NamedLocations") or []:
        add(bag, l.get("Name"), "named location")

    for r in d.get("Roles") or []:
        for m in r.get("Members") or []:
            add(bag, m.get("DisplayName"), "role member name")
            add(bag, m.get("UserPrincipalName"), "role member UPN")

    g = d.get("Groups") or {}
    for n in g.get("RoleAssignable") or []:
        add(bag, n, "role-assignable group")
    for dyn in g.get("Dynamic") or []:
        add(bag, dyn.get("Name"), "dynamic group")
        for lit in re.findall(r'"([^"]*)"', str(dyn.get("Rule") or "")):
            add(bag, lit, "membership rule literal")

    for m in d.get("AuthMethods") or []:
        for t in m.get("Targets") or []:
            add(bag, t, "auth method target", KEEP_TARGETS)

    for a in d.get("Applications") or []:
        add(bag, a.get("Name"), "app name")
        add(bag, a.get("AppId"), "app id")
        for c in a.get("Credentials") or []:
            add(bag, c.get("Name"), "credential name")

    i = d.get("Intune") or {}
    if i.get("Available"):
        for key in ("CompliancePolicies", "ConfigurationProfiles"):
            for p in i.get(key) or []:
                for asg in p.get("Assignments") or []:
                    add(bag, asg, "Intune assignment", KEEP_ASSIGNMENTS)


def survivors(d, bag=None):
    """Values that MUST still appear - proof it did not just redact everything."""
    out = {}
    for l in d.get("Licenses") or []:
        if l.get("Sku"):
            out.setdefault(str(l["Sku"]), set()).add("license SKU")
    for r in d.get("Roles") or []:
        if r.get("Role"):
            out.setdefault(str(r["Role"]), set()).add("built-in role name")
    for p in (d.get("ConditionalAccess") or {}).get("Policies") or []:
        if p.get("Name"):
            name = str(p["Name"])
            # A policy name that embeds a mapped identifier is deliberately
            # rewritten by the scrub, so verbatim survival is only expected
            # of names that embed none.
            if bag is not None and any(token_in(v, name) for v in bag):
                continue
            out.setdefault(name, set()).add("CA policy name (kept by design)")
    uc = d.get("UserCounts") or {}
    for k, v in uc.items():
        out.setdefault(str(v), set()).add("user count " + k)
    return out


def read_outputs(root):
    files = {}
    for dirpath, _dirs, names in os.walk(root):
        # history/ holds the deliberately-real archive; it is checked separately.
        if os.path.basename(dirpath) == "history":
            continue
        for n in names:
            p = os.path.join(dirpath, n)
            with open(p, "r", encoding="utf-8-sig", errors="replace") as fh:
                files[os.path.relpath(p, root)] = fh.read()
    return files


def main():
    out_root = sys.argv[1]

    bag = {}
    with open(os.path.join(REPO, "sample-data.json"), encoding="utf-8-sig") as fh:
        current = json.load(fh)
    harvest(current, bag)

    hist_dir = os.path.join(REPO, "sample-history")
    for n in sorted(os.listdir(hist_dir)):
        if n.endswith(".json"):
            with open(os.path.join(hist_dir, n), encoding="utf-8-sig") as fh:
                harvest(json.load(fh), bag)

    files = read_outputs(out_root)
    print("Scanning %d output file(s) for %d harvested identifier(s).\n"
          % (len(files), len(bag)))

    # Whole-token match: a short credential name like "main" is a real
    # identifier worth scanning for, but a substring test would flag the word
    # "domain" and bury a real leak in noise.
    hit = token_in

    leaks = []
    for value, whys in sorted(bag.items()):
        for name, text in sorted(files.items()):
            if hit(value, text):
                leaks.append((value, sorted(whys), name))

    keep = survivors(current, bag)
    missing = []
    for value, whys in sorted(keep.items()):
        if not any(hit(value, t) for t in files.values()):
            missing.append((value, sorted(whys)))

    if leaks:
        print("LEAKS (%d):" % len(leaks))
        for value, whys, name in leaks[:40]:
            print("  %-46s %-28s -> %s" % (value[:44], ",".join(whys)[:26], name))
        if len(leaks) > 40:
            print("  ... and %d more" % (len(leaks) - 40))
    else:
        print("LEAKS: none. No harvested identifier appears in any output file.")

    print("")
    if missing:
        print("MISSING SURVIVORS (%d) - anonymizer removed something it should keep:" % len(missing))
        for value, whys in missing:
            print("  %-46s %s" % (value[:44], ",".join(whys)))
    else:
        print("SURVIVORS: all %d expected values still present (SKUs, role names, "
              "policy names, counts)." % len(keep))

    print("")
    return 1 if (leaks or missing) else 0


if __name__ == "__main__":
    sys.exit(main())
