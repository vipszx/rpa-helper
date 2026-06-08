import os


def is_writable(filename):
    if not os.path.exists(filename):
        return False
    try:
        with open(filename, 'r+b') as f:
            pass
    except PermissionError:
        return False
    else:
        return True
