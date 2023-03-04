class Colors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKCYAN = '\033[96m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'
    RESET = '\033[0m'
    RED = '\033[91m'

def print_bold_warning(txt):
    print(Colors.BOLD + Colors.WARNING + txt + Colors.RESET)

def print_bold_blue(txt):
    print(Colors.BOLD + Colors.OKBLUE + txt + Colors.RESET)

def print_bold_red(txt):
    print(Colors.BOLD + Colors.RED + txt + Colors.RESET)

def print_bold_header(txt):
    print(Colors.BOLD + Colors.HEADER + txt + Colors.RESET)

def print_bold_green(txt):
    print(Colors.BOLD + Colors.OKGREEN + txt + Colors.RESET)