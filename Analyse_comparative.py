#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Analyse comparative de marché - Marché Lagune / Baobab Shop / Comptoir Ivoire
DA Python - L2 - séance 3

Programme réalisé strictement d'après les 8 consignes du LISEZ-MOI :
  1. Chargement   2. Nettoyage (+ journal)   3. Rapprochement   4. Comparatif (+ CSV)
  5. Synthèse     6. Panier                  7. Historique      8. Complexité

Utilisation :  python analyse_comparative.py
Les 5 fichiers de données peuvent être dans un dossier « donnees » à côté du script,
ou directement à côté du script. Dépendance externe : openpyxl  (pip install openpyxl)
"""

from __future__ import annotations

import bisect
import csv
import difflib
import heapq
import json
import re
import sys
import time
import unicodedata
from collections import Counter, defaultdict
from dataclasses import dataclass
from datetime import date, timedelta
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from statistics import mean, median
from typing import Dict, List, Optional, Tuple

try:
    import openpyxl
except ImportError:
    sys.exit("Le module « openpyxl » est requis (lecture du fichier Excel) :  pip install openpyxl")

# ----------------------------------------------------------------------------
# Paramètres issus du LISEZ-MOI
# ----------------------------------------------------------------------------
ML, BB, CI = "Marché Lagune", "Baobab Shop", "Comptoir Ivoire"
ENSEIGNES = (ML, BB, CI)
PANIER = {"P04": 1, "P12": 1, "P14": 2, "P15": 1, "P28": 1}
BF_DEBUT, BF_FIN = "2025-11-24", "2025-11-30"
DOSSIER_SORTIE = "sorties"


# ----------------------------------------------------------------------------
# Outils
# ----------------------------------------------------------------------------
def nettoyer(v) -> str:
    return "" if v is None else re.sub(r"\s+", " ", unicodedata.normalize("NFKC", str(v))).strip()


def sans_accents(t: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFD", t) if unicodedata.category(c) != "Mn")


def parse_prix(v) -> Optional[int]:
    """'102 900 FCFA', '279\u202f300', 109400, 'Prix sur demande', '' -> int ou None."""
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return int(round(v)) if v > 0 else None
    chiffres = re.sub(r"\D", "", nettoyer(v))
    return int(chiffres) if chiffres and int(chiffres) > 0 else None


def fcfa(n) -> str:
    return f"{int(round(n)):,}".replace(",", " ") + " FCFA"


def pct(x: float) -> str:
    return f"{x:.1f} %".replace(".", ",")


def afficher_tableau(entetes: List[str], lignes: List[list]) -> None:
    lignes = [[str(c) for c in l] for l in lignes]
    largeurs = [max(len(str(e)), *(len(l[i]) for l in lignes)) for i, e in enumerate(entetes)]
    print("  " + " | ".join(str(e).ljust(w) for e, w in zip(entetes, largeurs)))
    print("  " + "-+-".join("-" * w for w in largeurs))
    for l in lignes:
        print("  " + " | ".join(c.ljust(w) for c, w in zip(l, largeurs)))


# ----------------------------------------------------------------------------
# Journal de nettoyage (consigne 2)
# ----------------------------------------------------------------------------
JOURNAL: List[Tuple[str, str, str]] = []          # (source, problème, traitement)


def journaliser(source: str, probleme: str, traitement: str) -> None:
    JOURNAL.append((source, probleme, traitement))


def ecrire_journal(chemin: Path) -> None:
    lignes = ["JOURNAL DE NETTOYAGE", "=" * 20, f"{len(JOURNAL)} problème(s) repéré(s) et traité(s)", ""]
    for source in ["Catalogue", ML, BB, CI, "Historique", "Rapprochement", "Prix"]:
        entrees = [e for e in JOURNAL if e[0] == source]
        if entrees:
            lignes += [source.upper(), "-" * len(source)]
            for _, pb, tr in entrees:
                lignes += [f"* {pb}", f"    -> {tr}"]
            lignes.append("")
    chemin.write_text("\n".join(lignes), encoding="utf-8")


# ----------------------------------------------------------------------------
# Consigne 1 : une même structure de données pour les trois enseignes
# ----------------------------------------------------------------------------
@dataclass
class Offre:
    source: str
    ref: str
    libelle: str
    prix_brut: Optional[int]          # prix affiché
    frais_livraison: int = 0          # Baobab : delivery_fee
    remise_pct: float = 0.0           # Comptoir : remise
    stock: Optional[int] = None       # quantité connue (None = inconnue)
    rupture: bool = False
    ancien_prix: Optional[int] = None
    produit_id: Optional[str] = None
    prix_paye: Optional[int] = None   # ce que le client paie réellement

    def calculer_prix_paye(self) -> None:
        """Marché Lagune : prix affiché (livraison incluse) ; Baobab : amount + delivery_fee ;
        Comptoir : prix catalogue diminué de la remise."""
        if self.prix_brut is None:
            self.prix_paye = None
            return
        net = Decimal(self.prix_brut) * (Decimal(100) - Decimal(str(self.remise_pct))) / Decimal(100)
        self.prix_paye = int(net.quantize(Decimal("1"), rounding=ROUND_HALF_UP)) + self.frais_livraison

    @property
    def exploitable(self) -> bool:
        """Ni en rupture, ni sans prix : seule une offre exploitable peut être « la moins chère »."""
        return self.prix_paye is not None and not self.rupture

    @property
    def motif(self) -> str:
        return "en rupture" if self.rupture else ("sans prix exploitable" if self.prix_paye is None else "ok")


def charger_catalogue(chemin: Path) -> Dict[str, dict]:
    with open(chemin, encoding="utf-8-sig", newline="") as f:
        return {nettoyer(l["produit_id"]): {"nom": nettoyer(l["nom"]), "categorie": nettoyer(l["categorie"]),
                                            "marque": nettoyer(l["marque"])} for l in csv.DictReader(f)}


def charger_marche_lagune(chemin: Path) -> List[Offre]:
    with open(chemin, "rb") as f:
        if f.read(3) == b"\xef\xbb\xbf":
            journaliser(ML, "BOM UTF-8 au début du fichier", "lecture en utf-8-sig")
    offres, vues = [], {}
    doublons, texte_prix = [], 0
    with open(chemin, encoding="utf-8-sig", newline="") as f:
        for l in csv.DictReader(f, delimiter=";"):
            ref = nettoyer(l["reference"])
            signature = tuple(nettoyer(v) for v in l.values())
            if ref in vues:                                  # dict : détection des doublons en O(1)
                if vues[ref] == signature:
                    doublons.append(ref)
                continue
            vues[ref] = signature
            libelle = nettoyer(l["designation"])
            if l["designation"] != l["designation"].strip():
                journaliser(ML, f"{ref} : espaces superflus et minuscules « {l['designation']} »", "libellé nettoyé")
            prix_txt = nettoyer(l["prix"])
            prix = parse_prix(prix_txt)
            if prix is None:
                journaliser(ML, f"{ref} « {libelle} » : prix inutilisable (« {prix_txt or 'vide'} »)",
                            "offre conservée mais sans prix exploitable")
            else:
                texte_prix += 1
            ancien = parse_prix(l["ancien_prix"])
            dispo = sans_accents(nettoyer(l["disponibilite"]).lower())
            o = Offre(ML, ref, libelle, prix, ancien_prix=ancien, rupture=(dispo == "rupture"))
            o.calculer_prix_paye()
            offres.append(o)
    if doublons:
        journaliser(ML, f"lignes en double : {', '.join(doublons)}", "doublons supprimés")
    journaliser(ML, f"{texte_prix} prix écrits en texte (« 102 900 FCFA »)", "convertis en entiers")
    journaliser(ML, "note_clients avec virgule décimale", "champ non utilisé par l'analyse demandée")
    return offres


def charger_baobab(chemin: Path) -> List[Offre]:
    with open(chemin, encoding="utf-8") as f:
        items = json.load(f)["items"]
    offres, skus = [], set()
    for it in items:
        sku = nettoyer(it.get("sku"))
        if sku in skus:
            journaliser(BB, f"sku {sku} en double", "doublon supprimé")
            continue
        skus.add(sku)
        pr = it.get("pricing") or {}
        libelle = nettoyer(it.get("title"))
        prix = parse_prix(pr.get("amount"))
        if prix is None:
            journaliser(BB, f"{sku} « {libelle} » : amount = null", "offre conservée mais sans prix exploitable")
        stock = it.get("stock")
        o = Offre(BB, sku, libelle, prix, frais_livraison=parse_prix(pr.get("delivery_fee")) or 0,
                  stock=stock, rupture=(stock == 0))
        o.calculer_prix_paye()
        offres.append(o)
    journaliser(BB, "titres en anglais, ordre des mots différent, parfois sans marque", "traité au rapprochement")
    return offres


def charger_comptoir(chemin: Path) -> List[Offre]:
    ws = openpyxl.load_workbook(chemin, data_only=True).worksheets[0]
    lignes = list(ws.iter_rows(values_only=True))
    idx = next(i for i, l in enumerate(lignes[:15]) if l and nettoyer(l[0]).lower() == "code")
    if idx != 0:
        journaliser(CI, f"titre en ligne 1, en-têtes en ligne {idx + 1}", "en-têtes repérés automatiquement")
    entetes = [sans_accents(nettoyer(c)).lower() for c in lignes[idx]]
    col = lambda mot: next(i for i, e in enumerate(entetes) if mot in e)
    c_code, c_art, c_prix, c_rem, c_qte = col("code"), col("article"), col("prix"), col("remise"), col("quantite")
    offres = []
    for l in lignes[idx + 1:]:
        if not l or all(c is None for c in l):
            continue
        code, libelle, brut = nettoyer(l[c_code]), nettoyer(l[c_art]), l[c_prix]
        prix = parse_prix(brut)
        if isinstance(brut, str):
            journaliser(CI, f"{code} « {libelle} » : prix enregistré en texte « {brut} » (espace insécable)",
                        f"converti en nombre : {fcfa(prix)}")
        remise = float(l[c_rem]) if isinstance(l[c_rem], (int, float)) else 0.0
        qte = int(l[c_qte]) if isinstance(l[c_qte], (int, float)) else None
        o = Offre(CI, code, libelle, prix, remise_pct=remise, stock=qte, rupture=(qte == 0))
        o.calculer_prix_paye()
        offres.append(o)
    journaliser(CI, "noms abrégés ou fautifs (« Frigo », « Clim », « DD externe », « Samsun »)",
                "traité au rapprochement")
    return offres


# ----------------------------------------------------------------------------
# Historique : dict (enseigne, produit) -> série triée, interrogée par dichotomie (bisect)
# ----------------------------------------------------------------------------
class Serie:
    def __init__(self) -> None:
        self.dates: List[str] = []
        self.prix: List[int] = []

    def trier(self) -> None:
        if any(a > b for a, b in zip(self.dates, self.dates[1:])):
            paires = sorted(zip(self.dates, self.prix))
            self.dates, self.prix = [p[0] for p in paires], [p[1] for p in paires]

    def valeurs(self, debut: Optional[str] = None, fin: Optional[str] = None) -> List[int]:
        i = bisect.bisect_left(self.dates, debut) if debut else 0
        j = bisect.bisect_right(self.dates, fin) if fin else len(self.dates)
        return self.prix[i:j]

    def moyenne(self, debut: Optional[str] = None, fin: Optional[str] = None) -> Optional[float]:
        v = self.valeurs(debut, fin)
        return sum(v) / len(v) if v else None


def charger_historique(chemin: Path, catalogue: Dict[str, dict]):
    series: Dict[Tuple[str, str], Serie] = {}
    lignes: List[Tuple[str, str, str, int]] = []
    vues, rejets = set(), Counter()
    with open(chemin, encoding="utf-8-sig", newline="") as f:
        lecteur = csv.reader(f)
        next(lecteur)
        for l in lecteur:
            d, src, pid, px = (nettoyer(c) for c in l)
            src, prix = unicodedata.normalize("NFC", src), parse_prix(px)
            if src not in ENSEIGNES or pid not in catalogue or prix is None or not re.fullmatch(r"\d{4}-\d{2}-\d{2}", d):
                rejets["ligne invalide"] += 1
            elif (d, src, pid) in vues:
                rejets["doublon"] += 1
            else:
                vues.add((d, src, pid))
                s = series.setdefault((src, pid), Serie())
                s.dates.append(d)
                s.prix.append(prix)
                lignes.append((d, src, pid, prix))
    for s in series.values():
        s.trier()
    for motif, n in rejets.items():
        journaliser("Historique", f"{n} ligne(s) : {motif}", "lignes ignorées")
    if not rejets:
        journaliser("Historique", f"{len(lignes)} lignes contrôlées (dates, enseignes, produits, prix, doublons)",
                    "aucune anomalie, aucun traitement")
    return series, lignes


# ----------------------------------------------------------------------------
# Consigne 3 : rapprochement  libellé -> produit_id
# ----------------------------------------------------------------------------
UNITES = {"gb": "go", "go": "go", "tb": "to", "to": "to", "l": "l", "litre": "l", "litres": "l",
          "cv": "cv", "hp": "cv", "pouce": "pouces", "pouces": "pouces", "inch": "pouces"}
SYNONYMES = {"frigo": "refrigerateur", "fridge": "refrigerateur", "ventilo": "ventilateur", "clim": "climatiseur",
             "conditioner": "climatiseur", "blender": "mixeur", "iron": "fer", "kettle": "bouilloire",
             "monitor": "ecran", "mouse": "souris", "printer": "imprimante", "router": "routeur", "dd": "disque dur"}
MOTS_VIDES = {"a", "de", "du", "des", "la", "le", "les", "et", "sans", "fil", "sur", "pied", "wireless", "electrique",
              "electric", "steam", "air", "pc", "portable", "laptop", "ssd"}
RE_UNITE = re.compile(r"^(\d+(?:\.\d+)?)(go|to|l|cv|pouces)$")
RE_CODE = re.compile(r"^([a-z]+)(\d+[a-z]*)$")


def jetons(texte: str) -> List[str]:
    """Libellé -> mots normalisés (sans accents, 128GB = 128 Go, 1,5 L = 1.5l, synonymes, mots vides retirés)."""
    t = sans_accents(nettoyer(texte)).lower().replace('"', " pouces ")
    t = re.sub(r"micro[\s-]?ondes|microwave", "microondes", t)
    t = re.sub(r"(\d),(\d)", r"\1.\2", t)
    t = re.sub(r"[^a-z0-9.]+", " ", t)
    t = re.sub(r"(?<!\d)\.|\.(?!\d)", " ", t)
    t = re.sub(r"(\d+(?:\.\d+)?)\s*(gb|go|tb|to|litres?|l|cv|hp|pouces?|inch)\b",
               lambda m: m.group(1) + UNITES[m.group(2)], t)
    sortie: List[str] = []
    for mot in t.split():
        for m in SYNONYMES.get(mot, mot).split():
            if m not in MOTS_VIDES and m not in sortie:
                sortie.append(m)
    return sortie


class Rapprocheur:
    def __init__(self, catalogue: Dict[str, dict]) -> None:
        self.marques = {t for p in catalogue.values() for t in jetons(p["marque"])}
        self.jetons_cat = {pid: jetons(p["nom"]) for pid, p in catalogue.items()}
        self.index: Dict[str, set] = defaultdict(set)         # index inversé : mot -> {produit_id}
        for pid, toks in self.jetons_cat.items():
            for t in toks:
                self.index[t].add(pid)
        self.vocabulaire = [t for t in self.index if t.isalpha() and len(t) >= 4]

    def poids(self, t: str) -> int:
        if t in self.marques:
            return 3
        if RE_UNITE.match(t) or t.isdigit():
            return 2
        return 3 if (re.search(r"[a-z]", t) and re.search(r"\d", t)) else 1

    @staticmethod
    def _conflit(a: List[str], b: List[str]) -> bool:
        """Contradiction entre deux libellés : 128 Go / 256 Go, 1 CV / 1.5 CV, A15 / A25, i3 / i5."""
        def groupes(tokens):
            unites, codes = defaultdict(set), defaultdict(set)
            for t in tokens:
                m = RE_UNITE.match(t)
                if m:
                    unites[m.group(2)].add(m.group(1))
                elif RE_CODE.match(t):
                    codes[RE_CODE.match(t).group(1)].add(RE_CODE.match(t).group(2))
            return unites, codes
        return any(not ga[k] & gb[k] for ga, gb in zip(groupes(a), groupes(b)) for k in ga.keys() & gb.keys())

    def evaluer(self, tok: List[str], pid: str):
        tok_cat = self.jetons_cat[pid]
        if self._conflit(tok, tok_cat):
            return 0.0, []
        restants, w_ok, w_item, inconnus, fautes = set(tok_cat), 0, 0, [], []
        for t in tok:
            if t in restants:
                restants.discard(t)
                w_ok += self.poids(t)
                w_item += self.poids(t)
            else:
                inconnus.append(t)
        for t in inconnus:                                    # fautes de frappe : « samsun » ~ « samsung »
            proche = None
            if t.isalpha() and len(t) >= 4:
                for c in restants:
                    r = difflib.SequenceMatcher(None, t, c).ratio()
                    if c.isalpha() and len(c) >= 4 and r >= 0.85 and (proche is None or r > proche[0]):
                        proche = (r, c)
            if proche:
                restants.discard(proche[1])
                w_ok += self.poids(proche[1])
                w_item += self.poids(proche[1])
                fautes.append((t, proche[1]))
            else:
                w_item += self.poids(t)
        w_cat = sum(self.poids(t) for t in tok_cat)
        return (0.5 * (w_ok / w_item + w_ok / w_cat) if w_item and w_cat else 0.0), fautes

    def rapprocher(self, libelle: str):
        tok = jetons(libelle)
        candidats = set()
        for t in tok:
            if t in self.index:
                candidats |= self.index[t]
            elif t.isalpha() and len(t) >= 4:
                for v in difflib.get_close_matches(t, self.vocabulaire, n=2, cutoff=0.85):
                    candidats |= self.index[v]
        notes = sorted(((self.evaluer(tok, p), p) for p in candidats), key=lambda x: x[0][0], reverse=True)
        if not notes or notes[0][0][0] < 0.60:
            return None, []
        if len(notes) > 1 and notes[0][0][0] - notes[1][0][0] < 0.10:
            return None, []                                   # ambigu : on ne devine pas
        return notes[0][1], notes[0][0][1]


def rapprocher_offres(offres: List[Offre], catalogue: Dict[str, dict]) -> None:
    rap = Rapprocheur(catalogue)
    for o in offres:
        pid, fautes = rap.rapprocher(o.libelle)
        if pid is None:
            journaliser("Rapprochement", f"{o.source} {o.ref} « {o.libelle} » : aucun produit sûr", "article ignoré")
            continue
        o.produit_id = pid
        for faute, juste in fautes:
            journaliser("Rapprochement", f"{o.source} {o.ref} : « {faute} » pour « {juste} »",
                        f"faute tolérée, rapproché de {pid} - {catalogue[pid]['nom']}")
    vus: Dict[Tuple[str, str], Offre] = {}
    for o in offres:
        if o.produit_id:
            if (o.source, o.produit_id) in vus:
                journaliser("Rapprochement", f"{o.source} : deux articles pour {o.produit_id}", "premier conservé")
                o.produit_id = None
            else:
                vus[(o.source, o.produit_id)] = o


def corriger_prix_aberrants(offres: List[Offre], series: Dict[Tuple[str, str], Serie]) -> None:
    """Un prix dix fois (ou plus) trop grand/petit par rapport à l'historique récent est une erreur de saisie."""
    for o in offres:
        serie = series.get((o.source, o.produit_id))
        if not (o.prix_brut and serie and serie.dates):
            continue
        limite = (date.fromisoformat(serie.dates[-1]) - timedelta(days=30)).isoformat()
        ref = median(serie.valeurs(debut=limite))
        if 1 / 3 <= o.prix_brut / ref <= 3:
            continue
        ancien = o.prix_brut
        for facteur in (10, 100, 0.1, 0.01):
            if abs(ancien / facteur / ref - 1) <= 0.25:
                o.prix_brut = int(round(ancien / facteur))
                o.calculer_prix_paye()
                journaliser("Prix", f"{o.source} {o.ref} « {o.libelle} » : prix {fcfa(ancien)} aberrant "
                                    f"(historique récent : {fcfa(ref)})",
                            f"zéro en trop probable, remplacé par {fcfa(o.prix_brut)}")
                break
        else:
            o.prix_brut = None
            o.calculer_prix_paye()
            journaliser("Prix", f"{o.source} {o.ref} : prix {fcfa(ancien)} aberrant", "prix exclu")


# ----------------------------------------------------------------------------
# Consigne 4 : comparatif
# ----------------------------------------------------------------------------
def comparatif(offres: List[Offre], catalogue: Dict[str, dict]):
    index = {(o.source, o.produit_id): o for o in offres if o.produit_id}       # (enseigne, produit) -> offre
    details = []
    for pid in sorted(catalogue):
        prix = {s: index[(s, pid)].prix_paye for s in ENSEIGNES if (s, pid) in index and index[(s, pid)].exploitable}
        if len(prix) < 2:
            continue
        p_min, p_max = min(prix.values()), max(prix.values())
        details.append({"pid": pid, "nom": catalogue[pid]["nom"], "prix": prix, "min": p_min, "max": p_max,
                        "moins_chere": [s for s in ENSEIGNES if prix.get(s) == p_min],
                        "plus_chere": [s for s in ENSEIGNES if prix.get(s) == p_max],
                        "ecart": p_max - p_min, "ecart_pct": (p_max - p_min) / p_min * 100})
    return index, details


def exporter_comparatif(chemin: Path, details: List[dict]) -> None:
    with open(chemin, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["produit_id", "produit"] + [f"prix_{s}" for s in ENSEIGNES]
                   + ["moins_chere", "ecart_fcfa", "ecart_pct"])
        for d in details:
            w.writerow([d["pid"], d["nom"]] + [d["prix"].get(s, "") for s in ENSEIGNES]
                       + [" / ".join(d["moins_chere"]), d["ecart"], f"{d['ecart_pct']:.2f}".replace(".", ",")])


# ----------------------------------------------------------------------------
# Consigne 5 : synthèse
# ----------------------------------------------------------------------------
def synthese(details: List[dict]) -> None:
    victoires = Counter(s for d in details for s in d["moins_chere"])           # Counter : comptage en 1 passage
    ecarts = defaultdict(list)
    for d in details:
        for s, p in d["prix"].items():
            ecarts[s].append((p - d["min"], (p - d["min"]) / d["min"] * 100))
    print("\nCombien de fois chaque enseigne est la moins chère "
          f"(sur {len(details)} produits comparés ; un ex æquo compte pour chaque enseigne) :")
    afficher_tableau(["Enseigne", "Fois la moins chère"], [[s, victoires.get(s, 0)] for s in ENSEIGNES])
    print("\nÉcart moyen de chaque enseigne par rapport au prix le plus bas :")
    afficher_tableau(["Enseigne", "Écart moyen (FCFA)", "Écart moyen (%)", "Produits"],
                     [[s, fcfa(mean(e[0] for e in ecarts[s])), pct(mean(e[1] for e in ecarts[s])), len(ecarts[s])]
                      for s in ENSEIGNES])
    print("\nLes trois produits où il y a le plus à gagner à comparer (plus gros écart en FCFA) :")
    top = heapq.nlargest(3, details, key=lambda d: d["ecart"])                  # top 3 en O(n log 3)
    afficher_tableau(["Produit", "Moins cher", "Plus cher", "Écart", "Écart %"],
                     [[d["nom"], f"{' / '.join(d['moins_chere'])} ({fcfa(d['min'])})",
                       f"{' / '.join(d['plus_chere'])} ({fcfa(d['max'])})", fcfa(d["ecart"]), pct(d["ecart_pct"])]
                      for d in top])


# ----------------------------------------------------------------------------
# Consigne 6 : panier
# ----------------------------------------------------------------------------
def panier(index: Dict[Tuple[str, str], Offre], catalogue: Dict[str, dict]) -> None:
    def offre_utilisable(s: str, pid: str, qte: int):
        o = index.get((s, pid))
        if o is None:
            return None, "non référencé"
        if not o.exploitable:
            return None, o.motif
        if o.stock is not None and o.stock < qte:
            return None, f"stock insuffisant ({o.stock})"
        return o, "ok"

    print("\nPanier : 1 Tecno Spark 20, 1 Lenovo IdeaPad 3, 2 clés USB SanDisk, 1 disque dur Seagate, 1 micro-ondes Samsung")
    lignes, totaux = [], {s: 0 for s in ENSEIGNES}
    complet = {s: True for s in ENSEIGNES}
    total_mixte = 0
    for pid, qte in PANIER.items():
        ligne, options = [f"{qte} x {catalogue[pid]['nom']}"], []
        for s in ENSEIGNES:
            o, motif = offre_utilisable(s, pid, qte)
            if o:
                totaux[s] += o.prix_paye * qte
                options.append((o.prix_paye, s))
                ligne.append(fcfa(o.prix_paye * qte))
            else:
                complet[s] = False
                ligne.append(f"indisponible ({motif})")
        prix, s = min(options)
        total_mixte += prix * qte
        ligne.append(f"{s} : {fcfa(prix * qte)}")
        lignes.append(ligne)
    lignes.append(["TOTAL"] + [fcfa(totaux[s]) if complet[s] else "panier incomplet" for s in ENSEIGNES]
                  + [fcfa(total_mixte)])
    afficher_tableau(["Article"] + list(ENSEIGNES) + ["Au plus bas par article"], lignes)


# ----------------------------------------------------------------------------
# Consigne 7 : historique, semaine du Black Friday vs prix moyen de l'année
# ----------------------------------------------------------------------------
def historique(series: Dict[Tuple[str, str], Serie], catalogue: Dict[str, dict]) -> None:
    print(f"\nSemaine du Black Friday ({BF_DEBUT} au {BF_FIN}) comparée au prix moyen de l'année "
          "(chaque produit est comparé à sa propre moyenne, puis on fait la moyenne des écarts) :")
    lignes = []
    for s in ENSEIGNES:
        variations = []
        for pid in catalogue:
            serie = series.get((s, pid))
            if serie:
                bf, annee = serie.moyenne(BF_DEBUT, BF_FIN), serie.moyenne()
                if bf and annee:
                    variations.append((bf / annee - 1) * 100)
        v = mean(variations)
        lignes.append([s, f"{v:+.1f} %".replace(".", ","), f"{sum(1 for x in variations if x < 0)} / {len(variations)}",
                       "moins cher que la moyenne : bon moment pour acheter" if v < 0
                       else "plus cher que la moyenne : mieux vaut acheter à un autre moment"])
    afficher_tableau(["Enseigne", "Prix BF vs moyenne annuelle", "Produits moins chers au BF", "Conclusion"], lignes)


# ----------------------------------------------------------------------------
# Consigne 8 : complexité
# ----------------------------------------------------------------------------
def complexite(series, lignes_histo, duree_totale: float, duree_chargement: float) -> None:
    cles = sorted(series)
    requetes = [cles[i % len(cles)] for i in range(300)]
    t0 = time.perf_counter()
    for src, pid in requetes:                              # A : on reparcourt les 20 876 lignes à chaque requête
        somme = n = 0
        for d, s, p, x in lignes_histo:
            if s == src and p == pid and BF_DEBUT <= d <= BF_FIN:
                somme += x
                n += 1
    t_a = time.perf_counter() - t0
    t0 = time.perf_counter()
    for k in requetes:                                     # B : dict + bisect
        series[k].moyenne(BF_DEBUT, BF_FIN)
    t_b = time.perf_counter() - t0
    print("\nTemps d'exécution :")
    print(f"  - programme complet : {duree_totale * 1000:.0f} ms "
          f"(dont lecture + indexation de l'historique : {duree_chargement * 1000:.0f} ms)".replace(".", ","))
    print(f"  - 300 recherches de prix sur {len(lignes_histo)} lignes : "
          f"parcours de liste {t_a * 1000:.1f} ms contre dictionnaire + bisect {t_b * 1000:.2f} ms "
          f"(environ {round(t_a / t_b)} fois plus rapide)".replace(".", ","))
    print("\nChoix des structures de données :")
    for besoin, structure, pourquoi in [
        ("catalogue, offres par (enseigne, produit)", "dict", "accès direct en O(1) au lieu de parcourir une liste en O(n)"),
        ("doublons (références déjà vues)", "set / dict", "test d'appartenance en O(1), un seul passage O(n)"),
        ("rapprochement des libellés", "index inversé (dict mot -> produits)",
         "on ne compare qu'aux produits qui partagent un mot, pas aux 30"),
        ("historique (20 876 lignes)", "dict (enseigne, produit) -> listes triées par date",
         "recherche d'une période par dichotomie (bisect) en O(log n) au lieu de O(n)"),
        ("comptage des victoires", "Counter", "comptage en un seul passage"),
        ("3 plus gros écarts", "heapq.nlargest", "O(n log 3) au lieu de trier toute la liste en O(n log n)"),
    ]:
        print(f"  - {besoin} : {structure} -> {pourquoi}")


# ----------------------------------------------------------------------------
# Programme principal
# ----------------------------------------------------------------------------
def trouver_dossier_donnees() -> Path:
    ici = Path(__file__).resolve().parent
    for cand in (ici / "donnees", ici, Path.cwd() / "donnees", Path.cwd()):
        if (cand / "catalogue_reference.csv").exists():
            return cand
    sys.exit("Fichiers de données introuvables : placez-les dans un dossier « donnees » à côté du script, "
             "ou directement à côté du script.")


def main() -> None:
    try:
        sys.stdout.reconfigure(errors="replace")
    except AttributeError:
        pass
    debut = time.perf_counter()
    dossier = trouver_dossier_donnees()
    sortie = Path(__file__).resolve().parent / DOSSIER_SORTIE
    sortie.mkdir(exist_ok=True)

    # 1. Chargement (et 2. nettoyage consigné au fil de la lecture)
    catalogue = charger_catalogue(dossier / "catalogue_reference.csv")
    offres = (charger_marche_lagune(dossier / "marche_lagune.csv") + charger_baobab(dossier / "baobab_shop.json")
              + charger_comptoir(dossier / "comptoir_ivoire.xlsx"))
    for src in ENSEIGNES:
        rupt = [o.libelle for o in offres if o.source == src and o.rupture]
        journaliser(src, f"{len(rupt)} article(s) en rupture : {', '.join(rupt)}",
                    "conservés mais jamais retenus comme « le moins cher »")
    t0 = time.perf_counter()
    series, lignes_histo = charger_historique(dossier / "historique_prix.csv", catalogue)
    duree_chargement = time.perf_counter() - t0
    print("1. CHARGEMENT")
    afficher_tableau(["Enseigne", "Articles chargés"], [[s, sum(1 for o in offres if o.source == s)] for s in ENSEIGNES])
    print(f"  Historique : {len(lignes_histo)} lignes")

    # Le rapprochement et la correction des prix aberrants alimentent aussi le journal
    # (un prix n'est jugé aberrant qu'une fois le produit identifié) : on les exécute avant d'afficher 2 et 3.
    rapprocher_offres(offres, catalogue)
    corriger_prix_aberrants(offres, series)
    ecrire_journal(sortie / "journal_nettoyage.txt")

    print(f"\n2. NETTOYAGE : {len(JOURNAL)} problèmes repérés et traités -> {DOSSIER_SORTIE}/journal_nettoyage.txt")
    print("\n3. RAPPROCHEMENT")
    afficher_tableau(["Enseigne", "Articles rapprochés d'un produit_id"],
                     [[s, f"{sum(1 for o in offres if o.source == s and o.produit_id)} / "
                          f"{sum(1 for o in offres if o.source == s)}"] for s in ENSEIGNES])

    # 4. Comparatif
    index, details = comparatif(offres, catalogue)
    exporter_comparatif(sortie / "comparatif_prix.csv", details)
    print(f"\n4. COMPARATIF : {len(details)} produits disponibles dans au moins deux enseignes "
          f"-> {DOSSIER_SORTIE}/comparatif_prix.csv")
    afficher_tableau(["Produit"] + list(ENSEIGNES) + ["Moins chère", "Écart", "Écart %"],
                     [[d["nom"]] + [fcfa(d["prix"][s]) if s in d["prix"] else "-" for s in ENSEIGNES]
                      + [" / ".join(d["moins_chere"]), fcfa(d["ecart"]), pct(d["ecart_pct"])] for d in details])

    print("\n5. SYNTHÈSE")
    synthese(details)
    print("\n6. PANIER")
    panier(index, catalogue)
    print("\n7. HISTORIQUE")
    historique(series, catalogue)
    print("\n8. COMPLEXITÉ")
    complexite(series, lignes_histo, time.perf_counter() - debut, duree_chargement)


if __name__ == "__main__":
    try:
        main()
    except PermissionError as e:
        sys.exit(f"\nAccès refusé au fichier : {e.filename}\n"
                 "-> Fermez ce fichier s'il est ouvert dans Excel, et si le projet est dans OneDrive, faites un clic droit "
                 "sur le fichier > « Toujours conserver sur cet appareil » (ou déplacez le projet hors de OneDrive).")