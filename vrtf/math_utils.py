"""
Fonctions mathématiques : interpolation, polynômes, intégration.
Équivalents Python des fonctions VBA de ModuleFonctionsMath.bas.
"""

import math


def interp(x: float, xp, fp) -> float:
    """Interpolation linéaire par morceaux (remplace Interpolation_lineaire VBA)."""
    xp = list(xp)
    fp = list(fp)
    if x <= xp[0]:
        return fp[0]
    if x >= xp[-1]:
        return fp[-1]
    for i in range(len(xp) - 1):
        if xp[i] <= x <= xp[i + 1]:
            return linear_interp(xp[i], xp[i + 1], fp[i], fp[i + 1], x)
    return fp[-1]


def poly_cp(coeffs: list, T: float) -> float:
    """
    Évalue Cp(T) = a + b*T + c*T² + d*T³ + e*T⁴ + f*T⁵ + g*T⁶.
    coeffs = [a, b, c, d, e, f, g] (jusqu'à 7 termes).
    """
    result = 0.0
    for i, c in enumerate(coeffs):
        if c:
            result += c * (T ** i)
    return result


def poly_enthalpy(coeffs: list, T: float, T_ref: float = 273.0) -> float:
    """
    Intègre Cp(T) de T_ref à T : h = Σ c_j/j * (T^j - T_ref^j), j=1..n.
    Remplace Gethgaz VBA. Retourne [J/kg].
    """
    h = 0.0
    for i, c in enumerate(coeffs, start=1):
        if c:
            h += c / i * (T ** i - T_ref ** i)
    return h


def poly_integrate(coeffs: list, x1: float, x2: float) -> float:
    """Intégrale d'un polynôme (coeffs = [a0, a1, a2, ...]) de x1 à x2."""
    result = 0.0
    for i, c in enumerate(coeffs, start=1):
        if c:
            result += c / i * (x2 ** i - x1 ** i)
    return result


def bisect(f, a: float, b: float, tol: float = 1e-3, max_iter: int = 100) -> float:
    """
    Méthode de bisection : trouve x tel que f(x) = 0 dans [a, b].
    Remplace la dichotomie VBA (CalculTadiab, GasTemp).
    """
    fa = f(a)
    for _ in range(max_iter):
        m = (a + b) / 2.0
        fm = f(m)
        if abs(fm) <= tol or (b - a) / 2.0 < tol:
            return m
        if fa * fm < 0:
            b = m
        else:
            a = m
            fa = fm
    return (a + b) / 2.0


def linear_interp(x1: float, x2: float, y1: float, y2: float, x: float) -> float:
    """Interpolation linéaire entre deux points (remplace Interpolate VBA)."""
    return y1 + (y2 - y1) / (x2 - x1) * (x - x1)


def polint3(x: list, y: list) -> list:
    """
    Coefficients [a, b, c, d] du polynôme cubique passant par 4 points.
    Polynôme : f(t) = a + b*t + c*t^2 + d*t^3.
    Remplace Polint3 VBA (formule de Lagrange).
    """
    x1, x2, x3, x4 = x[0], x[1], x[2], x[3]
    y1, y2, y3, y4 = y[0], y[1], y[2], y[3]

    d1 = (x1 - x2) * (x1 - x3) * (x1 - x4)
    d2 = (x2 - x1) * (x2 - x3) * (x2 - x4)
    d3 = (x3 - x1) * (x3 - x2) * (x3 - x4)
    d4 = (x4 - x1) * (x4 - x2) * (x4 - x3)

    r1, r2, r3, r4 = y1 / d1, y2 / d2, y3 / d3, y4 / d4

    coef3 = r1 + r2 + r3 + r4

    coef2 = (r1 * (-x4 - x2 - x3)
           + r2 * (-x4 - x1 - x3)
           + r3 * (-x4 - x1 - x2)
           + r4 * (-x3 - x1 - x2))

    coef1 = (r1 * (x4 * (x2 + x3) + x2 * x3)
           + r2 * (x4 * (x1 + x3) + x1 * x3)
           + r3 * (x4 * (x1 + x2) + x1 * x2)
           + r4 * (x3 * (x1 + x2) + x1 * x2))

    coef0 = (r1 * (-x2 * x3 * x4)
           + r2 * (-x1 * x3 * x4)
           + r3 * (-x1 * x2 * x4)
           + r4 * (-x1 * x2 * x3))

    return [coef0, coef1, coef2, coef3]
