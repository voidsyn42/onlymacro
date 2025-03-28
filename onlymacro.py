#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Author: voidsyn42
# 📄 onlymacro.py

import sys
import json
import re

def vba_to_onlyoffice(vba_code: str):
    """Simple conversion of MsgBox from VBA to OnlyOffice-style macro JSON"""
    lines = vba_code.splitlines()
    for line in lines:
        match = re.search(r'MsgBox\s+"(.+?)"', line)
        if match:
            msg = match.group(1)
            return {
                "macros": [
                    {
                        "event": "onDocumentOpen",
                        "script": f"Api.ShowMessage('{msg}');"
                    }
                ]
            }
    return None

def main(vba_path):
    with open(vba_path, 'r') as f:
        vba_code = f.read()

    macro_json = vba_to_onlyoffice(vba_code)

    if macro_json:
        output_path = "macros.json"
        with open(output_path, 'w') as f:
            json.dump(macro_json, f, indent=2)
        print(f"[+] Macro converted and saved to {output_path}")
    else:
        print("[-] No supported macro found in the input file.")

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python3 onlymacro.py <vba_macro_file>")
        sys.exit(1)
    main(sys.argv[1])
