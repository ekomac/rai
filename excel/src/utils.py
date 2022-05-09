A_Z = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'


def get_name(index):
    """
    Returns the column name by index.
    """
    return A_Z[index]


def get_i(c: str):
    """
    Returns the column index by it's name.
    """
    return A_Z.index(c.upper())+1


A, B, C, D, E, F, G, H, I, J, K, L, M, N, O, P, \
    Q, R, S, T, U, V, W, X, Y, Z = [
        (lambda l: get_i(l))(letter) for letter in A_Z]
