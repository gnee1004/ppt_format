import sys
import os
import argparse
import olefile

OLE_SIG = b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1'  # Compound File Binary Format signature


def is_cfbf(path: str) -> bool:
    try:
        with open(path, 'rb') as f:
            sig = f.read(8)
        return sig == OLE_SIG
    except Exception:
        return False


def build_tree(ole: olefile.OleFileIO):
    entries = ole.listdir(streams=True, storages=True)
    tree = {}
    for entry in entries:
        node = tree
        for part in entry:
            node = node.setdefault(part, {})
    return tree


def print_tree(ole: olefile.OleFileIO, node: dict, prefix: str = "", path_prefix: tuple = ()):
    keys = sorted(node.keys())
    for i, name in enumerate(keys):
        is_last = (i == len(keys) - 1)
        branch = "└─ " if is_last else "├─ "
        spacer = "   " if is_last else "│  "
        child = node[name]
        full_path = path_prefix + (name,)

        # Determine if it's a stream (has size) or a storage (no size)
        label_type = "[storage]"
        size_info = ""
        try:
            size = ole.get_size(full_path)
            label_type = "[stream]"
            size_info = f"  ({size} bytes)"
        except Exception:
            # get_size raises if it's a storage
            pass

        print(prefix + branch + f"{name} {label_type}{size_info}")
        print_tree(ole, child, prefix + spacer, full_path)


def main():
    ap = argparse.ArgumentParser(description="Print OLE tree of a legacy .ppt (CFBF) file.")
    ap.add_argument("file", help="Path to .ppt")
    args = ap.parse_args()
    path = args.file

    if not os.path.exists(path):
        print(f"File not found: {path}")
        sys.exit(1)

    print(f"File: {path}")
    try:
        size = os.path.getsize(path)
        print(f"Size: {size} bytes")
    except Exception:
        pass

    print("CFBF signature:", "OK" if is_cfbf(path) else "NOT OLE/CFBF (probably .pptx)")

    try:
        ole = olefile.OleFileIO(path)
    except Exception as e:
        print(f"Failed to open as OLE: {e}")
        sys.exit(1)

    tree = build_tree(ole)
    print("\n== OLE Tree ==")
    print_tree(ole, tree)
    ole.close()


if __name__ == "__main__":
    main()
