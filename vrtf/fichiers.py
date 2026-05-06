"""
Sérialisation / désérialisation du projet four.
Remplace ModuleFichiers.bas (OpenProjectFile / SaveProjectFile VBA).

Le format VBA était un classeur Excel (.plf) avec une feuille-manifeste qui
listait les cellules à copier.  Ici on utilise du JSON : pas de dépendance
Excel, lisible par tout outil, compatible avec le reste du package Python.

L'extension .plf est conservée pour la compatibilité des noms de fichiers.
"""

import json
from pathlib import Path


def save_project(file_path: str | Path, data: dict) -> None:
    """
    Sauvegarde le dict projet dans un fichier JSON (.plf).
    Remplace SaveProjectFile VBA.

    file_path : chemin du fichier de destination (extension .plf recommandée)
    data      : dict python contenant tous les paramètres du projet (cfg, etc.)
    """
    path = Path(file_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)


def open_project(file_path: str | Path) -> dict:
    """
    Charge un fichier projet JSON (.plf) et retourne le dict.
    Remplace OpenProjectFile VBA.

    file_path : chemin du fichier .plf
    Retourne  : dict python avec les paramètres du projet
    Lève      : FileNotFoundError si le fichier n'existe pas
                ValueError si le contenu n'est pas un dict JSON valide
    """
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f"Fichier projet introuvable : {path}")
    with path.open("r", encoding="utf-8") as f:
        data = json.load(f)
    if not isinstance(data, dict):
        raise ValueError(f"Format invalide : le fichier doit contenir un objet JSON ({path})")
    return data
