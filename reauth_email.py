"""
One-time re-login for both sending accounts, so the app can see replies.

WHY THIS SCRIPT EXISTS
The original grant only asked for gmail.send. Detecting replies and bounces
needs gmail.metadata as well, and an OAuth token can't grow a scope it wasn't
issued with — it has to be re-consented. Running gmail_send.py directly does
work, but it prompts for an account key and a test recipient, which is a lot
of chances to get it wrong for something you do once. This just walks both
accounts in order and asks nothing.

Run it by double-clicking UJRA_BEJELENTKEZES.bat, or:
    python reauth_email.py
"""

import sys

import gmail_send


def main():
    print("=" * 62)
    print("  Gmail újra-bejelentkezés (egyszeri)")
    print("=" * 62)
    print()
    print("Ez azért kell, hogy az alkalmazás LÁSSA a válaszokat és a")
    print("visszapattanó leveleket. Enélkül az utánkövetés nem tud elindulni.")
    print()
    print("Amit a program kér: csak olvasási jog a fejlécekhez (gmail.metadata).")
    print("A leveleid TARTALMÁT nem tudja elolvasni, csak azt látja, hogy")
    print("érkezett-e üzenet a szálba és kitől.")
    print()

    # --force re-authorises every account regardless of what the token file
    # claims. Useful when Google has revoked a grant server-side but the
    # local JSON still looks intact.
    force = "--force" in sys.argv or "-f" in sys.argv

    todo = [(k, v) for k, v in gmail_send.ACCOUNTS.items()]
    # needs_reauth(), not missing_scopes(): the latter only reads the scope
    # list off disk, and a REVOKED token still lists its scopes. That is how
    # a dead landscale.agency@gmail.com grant got reported as "kesz" while
    # every send failed with invalid_grant.
    print("Fiokok ellenorzese (ez par masodperc)...")
    print()
    already = [] if force else [k for k, _ in todo if not gmail_send.needs_reauth(k)]
    if already:
        print(f"Ezek már készen vannak: {', '.join(already)}")
        print()

    failed = []
    for key, acct in todo:
        if not force and not gmail_send.needs_reauth(key):
            continue
        print("-" * 62)
        print(f"  Fiók: {acct['email']}")
        print("-" * 62)
        print()
        print(f"  Megnyílik a böngésző. FONTOS: ezzel a fiókkal jelentkezz be:")
        print(f"      {acct['email']}")
        print()
        print("  Ha 'Google hasn't verified this app' figyelmeztetést látsz:")
        print("      Advanced  ->  Go to ... (unsafe)   — ez a te saját appod.")
        print()
        input("  Nyomj ENTER-t a böngésző megnyitásához... ")
        try:
            gmail_send.get_service(key)
            print()
            print(f"  KESZ: {acct['email']} rendben.")
            print()
        except Exception as e:
            failed.append((acct["email"], str(e)[:200]))
            print()
            print(f"  HIBA ennél a fióknál: {str(e)[:200]}")
            print()

    print("=" * 62)
    remaining = [k for k in gmail_send.ACCOUNTS if gmail_send.needs_reauth(k)]
    if not remaining:
        print("  Mindkét fiók készen áll. Zárd be ezt az ablakot, és")
        print("  indítsd újra az email appot (START_EMAIL_HU.bat).")
    else:
        print(f"  Még hiányzik: {', '.join(remaining)}")
        for email, err in failed:
            print(f"    {email}: {err}")
    print("=" * 62)
    print()
    input("ENTER a bezáráshoz... ")


if __name__ == "__main__":
    main()
