import pandas as pd
from ldap3 import Server, Connection, ALL, NTLM
import argparse
import getpass
import sys

def get_user_by_name(first_name, last_name, conn, search_base):
    conn.search(
        search_base,
        f'(&(givenName={first_name})(sn={last_name}))',
        attributes=['cn', 'mail', 'givenName', 'sn']
    )
    return conn.entries

def print_entry(entry):
    print(f"Name: {entry.cn}")
    print(f"Email: {entry.mail}")
    print(f"First Name: {entry.givenName}")
    print(f"Last Name: {entry.sn}")
    print("-" * 40)

def main():
    parser = argparse.ArgumentParser(
        description="Query Active Directory for user details from an Excel file or by first and last name."
    )

    parser.add_argument("--server", required=True, help="Active Directory server address.")
    parser.add_argument("--user", required=True, help="Username for AD authentication.")
    parser.add_argument("--search-base", required=True, help="LDAP search base (e.g., 'dc=your,dc=domain,dc=com').")

    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--excel-file", help="Path to Excel file with 'First Name' and 'Last Name' columns.")
    group.add_argument("--fn", help="First name of the individual to search.")
    parser.add_argument("--ln", help="Last name of the individual to search (required with --fn).")

    args = parser.parse_args()

    if args.fn and not args.ln:
        parser.error("--ln (last name) is required when using --fn")

    password = getpass.getpass("Enter AD password: ")

    try:
        server = Server(args.server, get_info=ALL)
        conn = Connection(server, user=args.user, password=password, authentication=NTLM, auto_bind=True)
    except Exception as e:
        print(f"❌ Failed to connect to Active Directory: {e}")
        sys.exit(1)

    if args.excel_file:
        try:
            df = pd.read_excel(args.excel_file, usecols=['First Name', 'Last Name'])
        except Exception as e:
            print(f"❌ Failed to read Excel file: {e}")
            sys.exit(1)

        for _, row in df.iterrows():
            entries = get_user_by_name(row['First Name'], row['Last Name'], conn, args.search_base)
            if entries:
                for entry in entries:
                    print_entry(entry)
            else:
                print(f"⚠️ No entry found for {row['First Name']} {row['Last Name']}")
                print("-" * 40)

    elif args.fn and args.ln:
        entries = get_user_by_name(args.fn, args.ln, conn, args.search_base)
        if entries:
            for entry in entries:
                print_entry(entry)
        else:
            print(f"⚠️ No entry found for {args.fn} {args.ln}")

    conn.unbind()

if __name__ == "__main__":
    main()
