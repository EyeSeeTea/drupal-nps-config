def ansi(n, bold=False):
    "Return function that escapes text with ANSI color n"
    return lambda txt: f'\x1b[{n}{";1" if bold else ""}m{txt}\x1b[0m'

black, red, green, yellow, blue, magenta, cyan, white = map(ansi, range(30, 38))
blackB, redB, greenB, yellowB, blueB, magentaB, cyanB, whiteB = [
    ansi(i, bold=True) for i in range(30, 38)]
