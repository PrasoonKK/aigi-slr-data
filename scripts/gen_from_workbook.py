#!/usr/bin/env python3
"""Regenerate every number, table and bibliography entry used in the manuscript
and its supplement directly from the SLR workbook, so the paper can never drift
from the evidence base.

Usage (from the repository root):  python3 scripts/gen_from_workbook.py ["AIGI Workbook.xlsx"]
Outputs: generated/stats.tex, generated/*.tex tables, refs_corpus.bib
"""
import sys, re, collections, os, unicodedata, statistics
import openpyxl

WB = sys.argv[1] if len(sys.argv) > 1 else "AIGI Workbook.xlsx"
OUT = "generated"
os.makedirs(OUT, exist_ok=True)
wb = openpyxl.load_workbook(WB)

def rows(sheet, start, hdr_row):
    ws = wb[sheet]
    hdr = [c.value for c in ws[hdr_row]]
    out = []
    for r in ws.iter_rows(min_row=start, values_only=True):
        if r[0] and str(r[0]).startswith("R0"):
            out.append(dict(zip(hdr, r)))
    return out

M = {r["Record_ID"]: r for r in rows("04_Master_Records", 2, 1)}
TA = rows("05_Title_Abstract_Screening", 2, 1)
FT = rows("06_Full_Text_Screening", 2, 1)
AR = rows("10_Analysis_Ready", 4, 3)
TX = {r["Record_ID"]: r for r in rows("28_Taxonomy_Coding", 4, 3)}
QAws = wb["31_Quality_Appraisal"]
QA = {}
for r in QAws.iter_rows(min_row=12, values_only=True):
    if r[0] and str(r[0]).startswith("R0"):
        QA[r[0]] = dict(QA1=r[2], QA2=r[3], QA3=r[4], QA4=r[5], QA5=r[6], QA6=r[7], QA7=r[8], QA8=r[9], QAbasis=r[12])
RR = rows("32_Reported_Results", 2, 1)

N = len(AR)
Y = lambda r, h: r[h] == "YES"
pct = lambda n, d=None: f"{100*n/(d or N):.1f}"

def nocjk(s):
    s = re.sub(r"[\u2e80-\u9fff\uac00-\ud7af\uff00-\uffef\u3000-\u303f]+", " ", str(s))
    return re.sub(r"\s{2,}", " ", s).strip()

def tex(s):
    s = nocjk(str(s).split("; [")[0])
    s = s.replace("\\", "/").replace("&", r"\&").replace("%", r"\%").replace("_", r"\_").replace("#", r"\#").replace("$", r"\$")
    s = s.replace("~", r"\textasciitilde{}").replace("^", r"\^{}")
    return s

stats = {}
def S(name, val):
    stats[name] = val

# ---------------- PRISMA ----------------
S("Nidentified", 473); S("NdupRemoved", 91)
S("NscreenedTA", len(TA))
S("NexclTA", sum(1 for r in TA if r["Decision"] == "EXCLUDE"))
S("Nsought", len(FT))
S("Nnotretrieved", sum(1 for r in FT if r["Full_Text_Status"] == "Not retrieved"))
S("Nassessed", sum(1 for r in FT if r["Full_Text_Status"] == "RETRIEVED"))
exc = collections.Counter(r["Exclusion_Code"] for r in FT if r["Eligibility_Decision"] == "EXCLUDE")
S("NexclFT", sum(exc.values()))
for k, v in exc.items():
    S("NexclFT" + k.replace("FT-EC", "EC"), v)
S("Nincluded", sum(1 for r in FT if r["Eligibility_Decision"] == "INCLUDE"))
S("PctNotRetrieved", pct(stats["Nnotretrieved"], stats["Nsought"]))
assert stats["Nincluded"] == N
# TA exclusion codes (multi-coded)
tacodes = collections.Counter()
for r in TA:
    if r["Decision"] == "EXCLUDE":
        for c in re.split(r"[;,]", str(r["Exclusion_Code"])):
            c = c.strip()
            if c: tacodes[c] += 1
for k, v in tacodes.items():
    S("NTA" + k.replace("EC", "EC"), v)
S("NTAcodedEConeFive", sum(1 for r in TA if r["Decision"] == "EXCLUDE" and "EC1" in str(r["Exclusion_Code"]).split("; ") and "EC5" in str(r["Exclusion_Code"])))

# ---------------- corpus profile ----------------
S("N", N)
years = collections.Counter(int(float(r["Year"])) for r in AR)
for y in range(2020, 2027): S(f"Year{['Zero','One','Two','Three','Four','Five','Six'][y-2020]}", years.get(y, 0))
S("NyearTwentyFiveSix", years.get(2025, 0) + years.get(2026, 0))
S("PctyearTwentyFiveSix", pct(years.get(2025, 0) + years.get(2026, 0)))
S("Njournals", len(set(r["Journal"] for r in AR)))
task = collections.Counter(r["Primary_Task"] for r in AR)
for k, v in task.items(): S("Task" + re.sub(r"[^A-Za-z]", "", k.title()), v)
S("Nface", sum(Y(r, "Face_Specific") for r in AR)); S("Pctface", pct(stats["Nface"]))
S("NpendTA", sum(1 for r in TA if r["Validation_Status"] == "Pending")); S("NpendFT", sum(1 for r in FT if r["Validation_Status"] == "Pending")); S("NconfTAVeryHigh", sum(1 for r in TA if r["Confidence"] == "Very High"))
S("Nvalidated", sum(1 for r in AR if r["Validation_Status"] == "Validated"))
S("Npending", sum(1 for r in AR if r["Validation_Status"] == "Pending"))

# generators
for h, nm in [("GAN_Included", "GAN"), ("Diffusion_Included", "Diff"), ("Transformer_Autoregressive_Included", "AR"), ("Commercial_Models_Included", "Comm")]:
    S("N" + nm, sum(Y(r, h) for r in AR)); S("Pct" + nm, pct(stats["N" + nm]))
fam = collections.Counter(r["Generator_Families"] for r in AR)
S("FamGAN", fam["GAN"]); S("FamDiff", fam["DIFFUSION"]); S("FamGanDiff", fam["GAN+DIFFUSION"]); S("FamMulti", fam["MULTI-FAMILY"])
S("NmultiParadigm", fam["GAN+DIFFUSION"] + fam["MULTI-FAMILY"]); S("PctmultiParadigm", pct(stats["NmultiParadigm"]))

# generalization flags
GF = [("Unseen_Generator_Test", "Unseen"), ("Cross_Family_Test", "Family"), ("Cross_Dataset_Test", "Dataset"), ("Cross_Domain_Content_Test", "Domain"), ("Open_Set_Test", "Openset"), ("Continual_Adaptive_Test", "Continual")]
for h, nm in GF:
    S("N" + nm, sum(Y(r, h) for r in AR)); S("Pct" + nm, pct(stats["N" + nm]))
    S("Nno" + nm, sum(r[h] == "NO" for r in AR)); S("Pctno" + nm, pct(stats["Nno" + nm]))
zs = collections.Counter(r["Zero_Shot_or_Adapted"] for r in AR)
S("Nzeroshot", zs["ZERO-SHOT"]); S("Pctzeroshot", pct(zs["ZERO-SHOT"]))
S("Nadapted", zs["FINE-TUNED"] + zs["RETRAINED"] + zs["MIXED"] + zs["CALIBRATED"])
S("Ncalib", sum(1 for r in AR if r["Target_Data_Used_For_Calibration"] == "YES"))
S("Nnogen", sum(1 for r in AR if r["Generalization_Protocol"] == "NONE")); S("Pctnogen", pct(stats["Nnogen"]))
S("Nmultiple", sum(1 for r in AR if r["Generalization_Protocol"] == "MULTIPLE")); S("Pctmultiple", pct(stats["Nmultiple"]))
st = collections.Counter(r["Generalization_Strength"] for r in AR)
S("SVH", st["VERY HIGH"]); S("SH", st["HIGH"]); S("SM", st["MODERATE"]); S("SL", st["LOW"])
for k in ["VH", "H", "M", "L"]: S("PctS" + k, pct(stats["S" + k]))
S("Nclaim", sum(Y(r, "Authors_Claimed_Generalization") for r in AR))
S("Nsupport", sum(Y(r, "Evidence_Supports_Generalization") for r in AR))
mism = [r["Record_ID"] for r in AR if Y(r, "Authors_Claimed_Generalization") and r["Evidence_Supports_Generalization"] == "NO"]
S("Nmismatch", len(mism)); S("Pctmismatch", pct(len(mism)))
S("PctmismatchOfClaim", pct(len(mism), stats["Nclaim"]))

# ladder
L1 = [r for r in AR if Y(r, "Unseen_Generator_Test")]
L2 = [r for r in L1 if Y(r, "Cross_Family_Test") and Y(r, "Cross_Dataset_Test")]
L3 = [r for r in L2 if Y(r, "Robustness_Tested")]
L4 = [r for r in L3 if Y(r, "Open_Set_Test")]
L5 = [r for r in L4 if Y(r, "Continual_Adaptive_Test")]
for i, L in enumerate([L1, L2, L3, L4, L5], 1):
    S(f"Ladder{['','One','Two','Three','Four','Five'][i]}", len(L)); S(f"PctLadder{['','One','Two','Three','Four','Five'][i]}", pct(len(L)))
L3sm = [r for r in L2 if Y(r, "Social_Media_Postprocessing")]
S("NfamDataSocial", len(L3sm)); S("PctfamDataSocial", pct(len(L3sm)))
S("NgenNoRob", sum(1 for r in AR if Y(r, "Evidence_Supports_Generalization") and r["Robustness_Tested"] == "NO"))
S("NunseenNoOpen", sum(1 for r in AR if Y(r, "Unseen_Generator_Test") and r["Open_Set_Test"] == "NO"))
S("NfamNoData", sum(1 for r in AR if Y(r, "Cross_Family_Test") and r["Cross_Dataset_Test"] == "NO"))
S("NopenCont", sum(1 for r in AR if Y(r, "Open_Set_Test") and Y(r, "Continual_Adaptive_Test")))

# robustness
RB = [("Robustness_Tested", "Rob"), ("JPEG_Compression", "JPEG"), ("Blur", "Blur"), ("Resize", "Resize"), ("Noise", "Noise"), ("Social_Media_Postprocessing", "OSN"), ("Adversarial_Attack", "Adv")]
for h, nm in RB:
    S("N" + nm, sum(Y(r, h) for r in AR)); S("Pct" + nm, pct(stats["N" + nm]))
S("NnoOSN", N - stats["NOSN"]); S("PctnoOSN", pct(N - stats["NOSN"]))
S("NnoAdv", N - stats["NAdv"]); S("PctnoAdv", pct(N - stats["NAdv"]))
S("NgenOSN", sum(1 for r in AR if Y(r, "Evidence_Supports_Generalization") and Y(r, "Social_Media_Postprocessing")))
S("NgenAdv", sum(1 for r in AR if Y(r, "Evidence_Supports_Generalization") and Y(r, "Adversarial_Attack")))

# weakness (tested AND moderate/low)
weak = lambda h: sum(1 for r in AR if Y(r, h) and r["Generalization_Strength"] in ("MODERATE", "LOW"))
S("WeakFamily", weak("Cross_Family_Test")); S("WeakDataset", weak("Cross_Dataset_Test")); S("WeakDomain", weak("Cross_Domain_Content_Test"))
S("WeakRob", weak("Robustness_Tested")); S("WeakOpen", weak("Open_Set_Test"))

# methods / taxonomy
fams = ["HYBRID / MULTI-CUE", "SPATIAL / VISUAL ARTIFACT", "FREQUENCY / SPECTRAL", "SEMANTIC / FOUNDATION-MODEL", "RECONSTRUCTION / INVERSION", "SOURCE / FINGERPRINT FORENSIC", "NOISE / RESIDUAL FORENSIC", "GEOMETRY / PHYSICAL CUES"]
abbr_ = {"HYBRID / MULTI-CUE": "HYB", "SPATIAL / VISUAL ARTIFACT": "SPA", "FREQUENCY / SPECTRAL": "FRQ", "SEMANTIC / FOUNDATION-MODEL": "SEM", "RECONSTRUCTION / INVERSION": "REC", "SOURCE / FINGERPRINT FORENSIC": "SRC", "NOISE / RESIDUAL FORENSIC": "NOI", "GEOMETRY / PHYSICAL CUES": "GEO"}
short = {"HYBRID / MULTI-CUE": "Hybrid", "SPATIAL / VISUAL ARTIFACT": "Spatial", "FREQUENCY / SPECTRAL": "Freq", "SEMANTIC / FOUNDATION-MODEL": "Sem", "RECONSTRUCTION / INVERSION": "Rec", "SOURCE / FINGERPRINT FORENSIC": "Src", "NOISE / RESIDUAL FORENSIC": "Noise", "GEOMETRY / PHYSICAL CUES": "Geo"}
pf = collections.Counter(TX[r["Record_ID"]]["Primary_Evidence_Family"] for r in AR)
methodtasks = ("DETECTION", "DETECTION+ATTRIBUTION", "DETECTION+LOCALIZATION", "ATTRIBUTION")
METH = [r for r in AR if r["Primary_Task"] in methodtasks]
S("Nmethod", len(METH))
pfm = collections.Counter(TX[r["Record_ID"]]["Primary_Evidence_Family"] for r in METH)
anyfam = collections.Counter()
for r in AR:
    t = TX[r["Record_ID"]]
    s = {t["Primary_Evidence_Family"]} | {x.strip() for x in str(t["Secondary_Evidence_Families"]).split(";")}
    for f in s: anyfam[f] += 1
for f in fams:
    S("Fam" + short[f], pf[f]); S("PctFam" + short[f], pct(pf[f]))
    S("FamM" + short[f], pfm[f]); S("PctFamM" + short[f], pct(pfm[f], len(METH)))
    S("Any" + short[f], anyfam[f]); S("PctAny" + short[f], pct(anyfam[f]))
semcue = sum(1 for r in AR if str(r["Semantic_Features"]).strip().upper() not in ("NO", "", "NR", "N/A"))
S("Nsemcue", semcue); S("Pctsemcue", pct(semcue))
learn = collections.Counter(x.strip() for t in TX.values() for x in str(t["Learning_Strategy_Tags"]).split(";"))
for k, nm in [("ENSEMBLE / FUSION", "Ens"), ("FOUNDATION-MODEL TRANSFER", "FMT"), ("CONVENTIONAL DISCRIMINATIVE", "Conv"), ("OPEN-SET / UNCERTAINTY", "OSU"), ("CONTRASTIVE / METRIC LEARNING", "Contr"), ("DISENTANGLEMENT / INVARIANCE", "Dis"), ("CONTINUAL / ADAPTIVE", "Cont")]:
    S("Learn" + nm, learn[k]); S("PctLearn" + nm, pct(learn[k]))
# family x strength (high+very high rate)
for f in fams:
    rs = [r for r in AR if TX[r["Record_ID"]]["Primary_Evidence_Family"] == f]
    hv = sum(1 for r in rs if r["Generalization_Strength"] in ("HIGH", "VERY HIGH"))
    S("HV" + short[f], hv); S("PctHV" + short[f], pct(hv, len(rs)) if rs else "0.0")
S("Ncode", sum(Y(r, "Code_Available") for r in AR)); S("Pctcode", pct(stats["Ncode"]))
S("Ndata", sum(Y(r, "Dataset_Available") for r in AR)); S("Pctdata", pct(stats["Ndata"]))
S("Nablation", sum(Y(r, "Ablation_Study") for r in AR)); S("Pctablation", pct(stats["Nablation"]))
# QA distributions (QA1, QA3, QA8 are AI-scored and pending author validation; see sheet 31 column M)
QAK = ["QA1", "QA2", "QA3", "QA4", "QA5", "QA6", "QA7", "QA8"]
for k in QAK:
    assert all(isinstance(QA[r["Record_ID"]][k], (int, float)) for r in AR), k + " not fully scored"
    c = collections.Counter(QA[r["Record_ID"]][k] for r in AR)
    S(k + "One", c[1.0]); S(k + "Half", c[0.5]); S(k + "Zero", c[0.0])
    S(k + "Mean", f"{sum(QA[r['Record_ID']][k] for r in AR)/N:.2f}")
qtot = sorted(sum(QA[r["Record_ID"]][k] for k in QAK) for r in AR)
S("QATotalMean", f"{sum(qtot)/N:.2f}"); S("QATotalMedian", f"{statistics.median(qtot):.1f}")
S("QATotalMin", f"{qtot[0]:.1f}"); S("QATotalMax", f"{qtot[-1]:.1f}")
S("QATotalLowFour", sum(1 for q in qtot if q < 4)); S("QATotalHighSeven", sum(1 for q in qtot if q >= 7))
S("NqaEightFull", sum(1 for r in AR if str(QA[r["Record_ID"]]["QAbasis"]).find("full text") >= 0))

# ---------------- reported performance (sheet 32; values verbatim from extraction text) ----------------
for r in RR: r["Value_pct"] = float(r["Value_pct"])
med = lambda xs: statistics.median(xs)
def q(xs, p): return statistics.quantiles(xs, n=4)[0 if p == 1 else 2] if len(xs) > 1 else xs[0]
f1 = lambda x: f"{x:.1f}"
S("NresValues", len(RR)); S("NresStudies", len({r["Record_ID"] for r in RR}))
vals = lambda rid, st, ms: [r["Value_pct"] for r in RR if r["Record_ID"] == rid and r["Setting"] == st and r["Metric"] in ms]
PAIRS = {"Acc": [], "Rank": []}
for rid in sorted({r["Record_ID"] for r in RR if r["Pair_Eligible"] == "Y"}):
    i, o = vals(rid, "in-domain", ("ACC",)), vals(rid, "ood", ("ACC",))
    if i and o: PAIRS["Acc"].append((rid, "ACC", sum(i)/len(i), sum(o)/len(o), len(o)))
    for m in ("AP", "AUC"):
        i, o = vals(rid, "in-domain", (m,)), vals(rid, "ood", (m,))
        if i and o: PAIRS["Rank"].append((rid, m, sum(i)/len(i), sum(o)/len(o), len(o))); break
for k, P in PAIRS.items():
    d = [a - b for _, _, a, b, _ in P]
    S("Npair" + k, len(P)); S("MedID" + k, f1(med([p[2] for p in P]))); S("MedOOD" + k, f1(med([p[3] for p in P])))
    S("MedDrop" + k, f1(med(d))); S("QoneDrop" + k, f1(q(d, 1))); S("QthreeDrop" + k, f1(q(d, 3))); S("MaxDrop" + k, f1(max(d)))
    S("NdropTen" + k, sum(1 for x in d if x >= 10)); S("NdropTwenty" + k, sum(1 for x in d if x >= 20))
    S("MinOOD" + k, f1(min(p[3] for p in P)))
S("PctdropTenAcc", pct(stats["NdropTenAcc"], len(PAIRS["Acc"])))
# threshold transfer: OOD AP minus OOD ACC within the same study
thr = []
for rid in sorted({r["Record_ID"] for r in RR}):
    a, p = vals(rid, "ood", ("ACC",)), vals(rid, "ood", ("AP",))
    ta = {str(r["Target"]) for r in RR if r["Record_ID"] == rid and r["Setting"] == "ood" and r["Metric"] == "ACC"}
    tp = {str(r["Target"]) for r in RR if r["Record_ID"] == rid and r["Setting"] == "ood" and r["Metric"] == "AP"}
    if a and p and ta == tp: thr.append((rid, sum(p)/len(p) - sum(a)/len(a)))  # identical target sets only
S("Nthr", len(thr)); S("MedThr", f1(med([t for _, t in thr]))); S("MaxThr", f1(max(t for _, t in thr)))
S("NthrPos", sum(1 for _, t in thr if t > 0))
# in-domain-only studies (no generalization test) versus studies that test transfer
idonly = [r["Value_pct"] for r in RR if r["Setting"] == "in-domain" and r["Metric"] == "ACC"
          and next(a for a in AR if a["Record_ID"] == r["Record_ID"])["Generalization_Protocol"] == "NONE"]
S("NidOnly", len(idonly)); S("MedIDonly", f1(med(idonly))); S("MinIDonly", f1(min(idonly)))
oodacc = [r["Value_pct"] for r in RR if r["Setting"] == "ood" and r["Metric"] == "ACC"]
S("NoodAcc", len(oodacc)); S("MedOODallAcc", f1(med(oodacc))); S("MinOODallAcc", f1(min(oodacc)))
S("NoodAccBelowSeventy", sum(1 for v in oodacc if v < 70))
jp = [r["Value_pct"] for r in RR if r["Setting"] == "robustness" and r["Shift"] == "jpeg" and r["Metric"] == "ACC"]
S("NjpegVals", len(jp)); S("MedJpeg", f1(med(jp))); S("MinJpeg", f1(min(jp)))
osv = [r["Value_pct"] for r in RR if r["Setting"] == "open-set"]
S("NopenVals", len(osv)); S("MinOpen", f1(min(osv))); S("MaxOpen", f1(max(osv)))
with open(f"{OUT}/gap_points.tex", "w") as f:
    f.write("% AUTO-GENERATED: (in-domain, OOD) per study; same trained model, same metric\n")
    for k, P in PAIRS.items():
        f.write(f"\\newcommand{{\\GapPlot{k}}}{{\\addplot[only marks, mark=o, mark size=1.6pt] coordinates {{"
                + " ".join(f"({a:.2f},{b:.2f})" for _, _, a, b, _ in P) + "};}\n")
with open(f"{OUT}/gap_table.tex", "w") as f:
    f.write("% AUTO-GENERATED\n")
    for k in ("Acc", "Rank"):
        for rid, m, a, b, n in sorted(PAIRS[k], key=lambda p: p[2] - p[3], reverse=True):
            fam = abbr_[TX[rid]["Primary_Evidence_Family"]]
            f.write(f"\\cite{{{rid}}} & {fam} & {m} & {a:.1f} & {b:.1f} & {a-b:.1f} & {n} \\\\\n")
        if k == "Acc": f.write("\\midrule\n")
with open(f"{OUT}/results_table.tex", "w") as f:
    f.write("% AUTO-GENERATED from sheet 32_Reported_Results (each value is quoted verbatim in the workbook)\n")
    for r in RR:
        f.write(f"\\cite{{{r['Record_ID']}}} & {r['Setting']} & {tex(r['Shift'])} & {tex(r['Target'])} & {r['Metric']} & {r['Value_pct']:.2f} & {r['Stat']} \\\\\n")

# benchmark / metric usage (text-matched in extraction fields)
DSF = ("Training_Dataset", "Test_Dataset_s", "Training_Generator_s", "Test_Generator_s")
def mentions(pat, fields=DSF):
    rx = re.compile(pat, re.I)
    return [r["Record_ID"] for r in AR if any(rx.search(str(r[h] or "")) for h in fields)]
BENCH = {"GenImage": r"genimage", "ForenSynths": r"forensynth|cnn-?spot|wang et al\.? ?\(?2020|progan", "UFD": r"universalfake|ojha|ufd",
         "Chameleon": r"chameleon", "FFpp": r"faceforensics|ff\+\+", "CelebDF": r"celeb-?df", "DFDC": r"dfdc", "DiffForensics": r"diffusionforensics|dire dataset",
         "Synthbuster": r"synthbuster", "AIGCDB": r"aigcdetect", "CIFAKE": r"cifake", "WildRF": r"wildrf", "Midjourney": r"midjourney", "DALLE": r"dall", "FLUX": r"flux", "SDXL": r"sdxl|sd ?xl|stable diffusion xl"}
BL = {}
for k, pat in BENCH.items():
    BL[k] = mentions(pat); S("Bench" + k, len(BL[k])); S("PctBench" + k, pct(len(BL[k])))
for k, pat in {"ACC": r"\bacc|accuracy", "AP": r"\bm?ap\b|average precision", "AUC": r"auc|auroc", "EER": r"\beer\b", "Fone": r"\bf1|f-?score", "TPRFPR": r"tpr|fpr"}.items():
    n = len(mentions(pat, ("Metrics",))); S("Met" + k, n); S("PctMet" + k, pct(n))
S("NthreshFree", sum(1 for r in AR if re.search(r"auc|auroc|\bm?ap\b|average precision|\beer\b", str(r["Metrics"]), re.I)))
S("PctthreshFree", pct(stats["NthreshFree"]))
S("NaccOnly", sum(1 for r in AR if not re.search(r"auc|auroc|\bm?ap\b|average precision|\beer\b", str(r["Metrics"]), re.I)))
S("PctaccOnly", pct(stats["NaccOnly"]))

with open(f"{OUT}/stats.tex", "w") as f:
    f.write("% AUTO-GENERATED by scripts/gen_from_workbook.py -- do not edit by hand\n")
    W = dict(zip("0123456789", ["Zero","One","Two","Three","Four","Five","Six","Seven","Eight","Nine"]))
    for k, v in stats.items():
        k = "".join(W.get(ch, ch) for ch in k)
        assert re.fullmatch(r"[A-Za-z]+", k), k
        v = f"{v}\\%" if k.startswith("Pct") else v
        f.write(f"\\newcommand{{\\{k}}}{{{v}\\xspace}}\n")

# ---------------- citation lists ----------------
def cites(lst): return "\\cite{" + ",".join(lst) + "}"
lists = {
    "CiteMismatch": mism,
    "CiteOSN": [r["Record_ID"] for r in AR if Y(r, "Social_Media_Postprocessing")],
    "CiteAdv": [r["Record_ID"] for r in AR if Y(r, "Adversarial_Attack")],
    "CiteOpen": [r["Record_ID"] for r in AR if Y(r, "Open_Set_Test")],
    "CiteCont": [r["Record_ID"] for r in AR if Y(r, "Continual_Adaptive_Test")],
    "CiteVH": [r["Record_ID"] for r in AR if r["Generalization_Strength"] == "VERY HIGH"],
    "CiteLadderFour": [r["Record_ID"] for r in L4],
    "CiteNoGen": [r["Record_ID"] for r in AR if r["Generalization_Protocol"] == "NONE"],
    "CiteBench": [r["Record_ID"] for r in AR if r["Primary_Task"] == "BENCHMARK"],
    "CiteChameleon": BL["Chameleon"], "CiteCIFAKE": BL["CIFAKE"], "CiteGenImage": BL["GenImage"],
    "CiteRobStudy": [r["Record_ID"] for r in AR if r["Primary_Task"] == "ROBUSTNESS"],
    "CiteAttr": [r["Record_ID"] for r in AR if "ATTRIBUTION" in r["Primary_Task"]],
    "CiteFace": [r["Record_ID"] for r in AR if Y(r, "Face_Specific")],
    "CiteCalib": [r["Record_ID"] for r in AR if r["Target_Data_Used_For_Calibration"] == "YES"],
}
for f in fams:
    lists["CiteFam" + short[f]] = [r["Record_ID"] for r in AR if TX[r["Record_ID"]]["Primary_Evidence_Family"] == f]
with open(f"{OUT}/citelists.tex", "w") as f:
    f.write("% AUTO-GENERATED citation groups\n")
    for k, v in lists.items():
        f.write(f"\\newcommand{{\\{k}}}{{{cites(sorted(v))}}}\n")

# ---------------- per-study evidence table (supplement) ----------------
abbr = abbr_
taskab = {"DETECTION": "Det", "ATTRIBUTION": "Attr", "DETECTION+ATTRIBUTION": "Det+Attr", "DETECTION+LOCALIZATION": "Det+Loc", "BENCHMARK": "Bench", "ROBUSTNESS": "Rob", "OTHER": "Other"}
famab = {"GAN": "G", "DIFFUSION": "D", "GAN+DIFFUSION": "G+D", "MULTI-FAMILY": "Multi"}
mark = lambda r, h: r"\cmark" if r[h] == "YES" else (r"?" if r[h] in ("UNCLEAR",) else r"--")
with open(f"{OUT}/study_table.tex", "w") as f:
    for r in sorted(AR, key=lambda r: r["Record_ID"]):
        t = TX[r["Record_ID"]]; q = QA[r["Record_ID"]]
        name = tex(r["Title"].split(":")[0][:46])
        f.write(" & ".join([
            f"\\cite{{{r['Record_ID']}}} {r['Record_ID']}", str(int(float(r["Year"]))), name, taskab[r["Primary_Task"]], abbr[t["Primary_Evidence_Family"]], famab[r["Generator_Families"]],
            mark(r, "Unseen_Generator_Test"), mark(r, "Cross_Family_Test"), mark(r, "Cross_Dataset_Test"), mark(r, "Cross_Domain_Content_Test"), mark(r, "Open_Set_Test"), mark(r, "Continual_Adaptive_Test"),
            mark(r, "Robustness_Tested"), mark(r, "Social_Media_Postprocessing"), mark(r, "Adversarial_Attack"),
            {"VERY HIGH": "VH", "HIGH": "H", "MODERATE": "M", "LOW": "L"}[r["Generalization_Strength"]],
            f"{sum(q[k] for k in QAK):.1f}"]) + r" \\" + "\n")

# ---------------- bibliography (corpus + secondary) ----------------
def bibauthors(a):
    out = []
    for p in str(a).split(";"):
        p = p.strip()
        if not p: continue
        m = re.match(r"(.+?)\s+((?:[A-Z][a-z]?\.[\s\-]?)+)$", p)
        if m: out.append(f"{m.group(1)}, {m.group(2)}")
        else: out.append(p)
    return " and ".join(out) if out else "Anonymous"
def bibesc(s):
    s = nocjk(str(s).split("; [")[0])
    s = s.replace("&", r"\&").replace("%", r"\%").replace("_", r"\_").replace("#", r"\#")
    return s
inc_ids = sorted(r["Record_ID"] for r in AR)
secondary = ["R0030", "R0033", "R0169", "R0357"]
import json as _json
CROSSREF = _json.load(open("scripts/crossref_cache.json")) if os.path.exists("scripts/crossref_cache.json") else {}
with open("refs_corpus.bib", "w") as f:
    f.write("% AUTO-GENERATED from 04_Master_Records (Scopus export). Keys = SLR Record_IDs.\n")
    for k in inc_ids + secondary:
        m = M[k]
        doi = str(m["DOI"]).replace("\\", "")
        # volume, issue and pages come from scripts/crossref_cache.json (fetched from Crossref by DOI); the year follows Crossref when it differs
        cx = CROSSREF.get(k, {}); yr = int(float(m['Year'])); yr = cx.get("year") or yr
        extra = ""
        if cx.get("volume"): extra += f"  volume  = {{{cx['volume']}}},\n"
        if cx.get("issue"): extra += f"  number  = {{{cx['issue']}}},\n"
        pg = cx.get("page")
        if pg: extra += f"  pages   = {{{str(pg).replace('-', '--').replace('----', '--')}}},\n"
        elif cx.get("article"): extra += f"  note    = {{Art. no. {cx['article']}}},\n"
        f.write(f"@article{{{k},\n  author  = {{{bibauthors(m['Authors'])}}},\n  title   = {{{{{bibesc(m['Title'])}}}}},\n  journal = {{{bibesc(m['Journal'])}}},\n{extra}  year    = {{{yr}}},\n  doi     = {{{doi}}}\n}}\n\n")
print("stats", len(stats), "N", N)

# ---------------- full-text exclusions & not-retrieved (supplement) ----------------
with open(f"{OUT}/ft_excluded.tex", "w") as f:
    f.write("% AUTO-GENERATED\n")
    for r in FT:
        if r["Eligibility_Decision"] == "EXCLUDE":
            m = M.get(r["Record_ID"], {})
            f.write(f"{r['Record_ID']} & {tex(str(m.get('Title') or r.get('Title'))[:80])} & {tex(r['Exclusion_Code'])} \\\\\n")
with open(f"{OUT}/not_retrieved.tex", "w") as f:
    f.write("% AUTO-GENERATED\n")
    for r in FT:
        if r["Full_Text_Status"] == "Not retrieved":
            m = M.get(r["Record_ID"], {})
            f.write(f"{r['Record_ID']} & {tex(str(m.get('Title') or r.get('Title'))[:95])} & {int(float(m.get('Year') or 0))} \\\\\n")

# ---------------- taxonomy figures (Section 4) ----------------
FAMS = [("SPATIAL / VISUAL ARTIFACT", "Spatial"), ("FREQUENCY / SPECTRAL", "Freq"), ("NOISE / RESIDUAL FORENSIC", "Noise"),
        ("SOURCE / FINGERPRINT FORENSIC", "Src"), ("SEMANTIC / FOUNDATION-MODEL", "Sem"), ("RECONSTRUCTION / INVERSION", "Rec"),
        ("GEOMETRY / PHYSICAL CUES", "Geo"), ("HYBRID / MULTI-CUE", "Hybrid")]
LEARN = [("CONVENTIONAL DISCRIMINATIVE", "Conventional"), ("FOUNDATION-MODEL TRANSFER", "Foundation transfer"), ("CONTRASTIVE / METRIC LEARNING", "Contrastive"),
         ("DISENTANGLEMENT / INVARIANCE", "Disentanglement"), ("ENSEMBLE / FUSION", "Ensemble / fusion"), ("OPEN-SET / UNCERTAINTY", "Open-set"), ("CONTINUAL / ADAPTIVE", "Continual")]
SUBS = [("TEXTURE / LOCAL ARTIFACT", "Texture / local"), ("PATCH / REGION CONSISTENCY", "Patch consistency"), ("FOURIER / DCT SPECTRAL CUES", "Fourier / DCT"),
        ("WAVELET / MULTI-BAND", "Wavelet"), ("NOISE RESIDUAL / CAMERA-LIKE TRACES", "Noise residual"), ("GENERATOR / SOURCE FINGERPRINT", "Fingerprint"),
        ("CLIP / VLM REPRESENTATION", "CLIP / VLM"), ("PROMPT / CONCEPT ALIGNMENT", "Prompt alignment"), ("RECONSTRUCTION CONSISTENCY", "Reconstruction"),
        ("DIFFUSION TRAJECTORY / LATENT CUES", "Diffusion latent"), ("PHYSICAL / ISP CONSISTENCY", "Physical / ISP"), ("EXPLICIT MULTI-CUE FUSION", "Explicit fusion")]
def _tags(s):
    return [x.strip() for x in str(s).split(";") if x.strip() and x.strip() != "NONE"]
TXL = list(TX.values())
tl = []
def T_(name, val): tl.append(f"\\newcommand{{\\{name}}}{{{val}\\xspace}}")
conf = collections.Counter(t["Coding_Confidence"] for t in TXL)
T_("TaxConfHigh", conf.get("HIGH", 0)); T_("TaxConfMod", conf.get("MODERATE", 0)); T_("TaxConfLow", conf.get("LOW", 0))
T_("TaxStageFive", sum(1 for t in TXL if str(t["Coding_Notes"]).startswith("Stage 5")))
ns = collections.Counter(min(len(_tags(t["Secondary_Evidence_Families"])), 3) for t in TXL)
for k, nm in enumerate(["Zero", "One", "Two", "ThreePlus"]): T_(f"TaxSec{nm}", ns.get(k, 0))
nl = collections.Counter(len(_tags(t["Learning_Strategy_Tags"])) for t in TXL)
for k, nm in enumerate(["Zero", "One", "Two", "Three"]): T_(f"TaxLearn{nm}", nl.get(k, 0))
subc = collections.Counter(t["Primary_Subtype"] for t in TXL)
tl.append("\\newcommand{\\TaxSubSym}{" + ",".join(l.replace(" / ", "/") for _, l in SUBS[::-1]) + "}")
coords = " ".join(f"({subc.get(k, 0)},{l.replace(' / ', '/')})" for k, l in SUBS[::-1])
tl.append("\\newcommand{\\TaxSubCoords}{" + coords + "}")
cells = []
for fi, (fk, _) in enumerate(FAMS):
    for li, (lk, _) in enumerate(LEARN):
        n = sum(1 for t in TXL if t["Primary_Evidence_Family"] == fk and lk in _tags(t["Learning_Strategy_Tags"]))
        cells.append(f"\\taxcell{{{fi}}}{{{li}}}{{{n}}}")
tl.append("\\newcommand{\\TaxCells}{" + "".join(cells) + "}")
LSH = ["Conv", "FMT", "Contr", "Dis", "Ens", "OSU", "Cont"]
for fi, (fk, fs) in enumerate(FAMS):
    for li, (lk, _) in enumerate(LEARN):
        T_(f"TaxCell{fs}{LSH[li]}", sum(1 for t in TXL if t["Primary_Evidence_Family"] == fk and lk in _tags(t["Learning_Strategy_Tags"])))
T_("TaxSubTexture", subc.get("TEXTURE / LOCAL ARTIFACT", 0))
T_("TaxModSecZero", sum(1 for t in TXL if t["Coding_Confidence"] == "MODERATE" and not _tags(t["Secondary_Evidence_Families"])))
T_("TaxModSecTwoPlus", sum(1 for t in TXL if t["Coding_Confidence"] == "MODERATE" and len(_tags(t["Secondary_Evidence_Families"])) >= 2))

_sym = ",".join(l.replace(" / ", "/") for _, l in SUBS[::-1])
tl.append("\\newcommand{\\TaxSubPlot}{\\begin{tikzpicture}\\let\\xspace\\empty\\begin{axis}[xbar, width=\\columnwidth, height=7.2cm, bar width=7pt, symbolic y coords={" + _sym + "}, ytick=data, y dir=normal, xmin=0, xmax=44, xlabel={Studies (primary subtype)}, xlabel style={font=\\footnotesize}, tick label style={font=\\footnotesize}, nodes near coords, nodes near coords style={font=\\scriptsize}, axis lines*=left, enlarge y limits=0.06]\\addplot[fill=cB!70, draw=cA] coordinates {" + coords + "};\\end{axis}\\end{tikzpicture}}")
tl.append("\\newcommand{\\TaxMaxCell}{" + str(max(int(c.split('}{')[2].rstrip('}')) for c in cells)) + "}")
with open(f"{OUT}/taxonomy_data.tex", "w") as f:
    f.write("% AUTO-GENERATED by scripts/gen_from_workbook.py -- do not edit by hand\n" + "\n".join(tl) + "\n")

# ---------------- Section 5 figures: corpus profile and RQ1 ----------------
import re as _re
rl = []
def R_(name, val): rl.append(f"\\newcommand{{\\{name}}}{{{val}\\xspace}}")
PERIODS = [("a", "2020--23", lambda y: y <= 2023), ("b", "2024", lambda y: y == 2024), ("c", "2025", lambda y: y == 2025), ("d", "2026", lambda y: y == 2026)]
def per_rows(p):
    return [r for r in AR if p[2](int(float(r["Year"])))]
for p in PERIODS:
    R_("PerN" + p[0], len(per_rows(p)))
    R_("PerFace" + p[0], sum(1 for r in per_rows(p) if Y(r, "Face_Specific")))
journ = collections.Counter(r["Journal"] for r in AR)
R_("NjournalsOnce", sum(1 for v in journ.values() if v == 1)); R_("TopJournalN", max(journ.values())); R_("PctTopJournal", pct(max(journ.values())))
R_("NjournalsThree", sum(1 for v in journ.values() if v >= 3)); R_("NstudiesInThree", sum(v for v in journ.values() if v >= 3))
JSHORT = {"Neurocomputing": "Neurocomputing", "Knowledge-Based Systems": "Knowl.-Based Syst.", "Pattern Recognition Letters": "Pattern Recognit. Lett.",
          "Computer Vision and Image Understanding": "Comput. Vis. Image Underst.", "Applied Sciences": "Appl. Sci.", "IEEE Access": "IEEE Access",
          "IEEE Transactions on Circuits and Systems for Video Technology": "IEEE Trans. CSVT", "Expert Systems With Applications": "Expert Syst. Appl.",
          "Journal of Imaging": "J. Imaging", "IEEE Signal Processing Letters": "IEEE Signal Process. Lett."}
def hbar(pairs, xlabel, xmax, height, bw="7pt"):
    sym = ",".join(l for l, _ in pairs[::-1]); co = " ".join(f"({v},{l})" for l, v in pairs[::-1])
    _w = max(len(l) for l, _ in pairs) * 0.155 + 0.9
    return ("\\begin{tikzpicture}\\let\\xspace\\empty\\begin{axis}[xbar, width=\\columnwidth-" + f"{_w:.1f}cm" + ", height=" + height + ", bar width=" + bw + ", symbolic y coords={" + sym +
            "}, ytick=data, y dir=normal, xmin=0, xmax=" + str(xmax) + ", xlabel={" + xlabel + "}, xlabel style={font=\\footnotesize}, tick label style={font=\\footnotesize}, nodes near coords, nodes near coords style={font=\\scriptsize}, axis lines*=left, enlarge y limits=0.07]\\addplot[fill=cB!70, draw=cA] coordinates {" + co + "};\\end{axis}\\end{tikzpicture}")
jt = sorted(((JSHORT.get(k, k), v) for k, v in journ.items() if v >= 3), key=lambda x: (-x[1], x[0]))
rl.append("\\newcommand{\\JournalPlot}{" + hbar(jt, "Included studies", 9, "4.6cm") + "}")
BB = [("ViT/Swin transformer", r"\bViT|Swin|[Tt]ransformer|DeiT|BEiT"), ("ResNet / ResNeXt", r"ResNet|ResNeXt"), ("CLIP", r"\bCLIP\b"),
      ("Custom or lightweight CNN", r"[Cc]ustom|[Ll]ightweight"), ("EfficientNet", r"EfficientNet"), ("DINO", r"DINO"),
      ("MLLM/BLIP/Mamba", r"Mamba|BLIP|LLaVA|MLLM|Qwen|LLM|VLM"), ("Xception", r"Xception"), ("MobileNet", r"MobileNet"), ("VGG", r"VGG")]
bbc = [(l, sum(1 for r in AR if _re.search(rx, str(r["Backbone_Architecture"])))) for l, rx in BB]
for l, v in bbc: R_("BB" + _re.sub(r"[^A-Za-z]", "", l), v)
rl.append("\\newcommand{\\BackbonePlot}{" + hbar(sorted(bbc, key=lambda x: -x[1]), "Studies naming the architecture", 40, "5.6cm") + "}")
def stacked_cols(series, labels, ylabel, ymax, legend_cols):
    sym = ",".join(k for k, _, _ in labels)
    body = "".join(f"\\addplot[fill={fc}, draw=cA] coordinates {{{' '.join(f'({k},{v})' for (k, _, _), v in zip(labels, vals))}}};" for _, fc, vals in series)
    ticks = ",".join("{" + t + "\\\\$n{=}" + "\\PerN" + k + "$}" for k, t, _ in labels)
    leg = ",".join(n for n, _, _ in series)
    return ("\\begin{tikzpicture}\\let\\xspace\\empty\\begin{axis}[ybar stacked, width=\\columnwidth, height=5.4cm, bar width=22pt, symbolic x coords={" + sym + "}, xtick=data, xticklabels={" + ticks +
            "}, xticklabel style={align=center, font=\\scriptsize}, ymin=0, ymax=" + str(ymax) + ", ylabel={" + ylabel + "}, label style={font=\\footnotesize}, tick label style={font=\\footnotesize}, enlarge x limits=0.15, axis lines*=left, legend style={font=\\scriptsize, draw=none, at={(0.5,1.03)}, anchor=south, legend columns=" + str(legend_cols) + ", /tikz/every even column/.append style={column sep=3pt}}]" + body + "\\legend{" + leg + "}\\end{axis}\\end{tikzpicture}")
shade = ["cD", "cB!30", "cB!60", "cA", "cC!35", "cC!70"]
def share(rows_fn, cats):
    out = []
    for ci, (nm, test) in enumerate(cats):
        out.append((nm, shade[ci], [round(100 * sum(1 for r in per_rows(p) if test(r)) / max(1, len(per_rows(p))), 1) for p in PERIODS]))
    return out
GEN = [("GAN only", lambda r: r["Generator_Families"] == "GAN"), ("Diffusion only", lambda r: r["Generator_Families"] == "DIFFUSION"),
       ("GAN + diffusion", lambda r: r["Generator_Families"] == "GAN+DIFFUSION"), ("Multi-family", lambda r: r["Generator_Families"] == "MULTI-FAMILY")]
labs = [(p[0], p[1], None) for p in PERIODS]
rl.append("\\newcommand{\\PerGenPlot}{" + stacked_cols(share(None, GEN), labs, "Share of studies (\\%)", 100, 2) + "}")
for nm, test in GEN:
    for p in PERIODS: R_("PerGen" + _re.sub(r"[^A-Za-z]", "", nm) + p[0], sum(1 for r in per_rows(p) if test(r)))
pfam = lambda r: TX[r["Record_ID"]]["Primary_Evidence_Family"]
FAMC = [("Spatial", lambda r: pfam(r).startswith("SPATIAL")), ("Frequency", lambda r: pfam(r).startswith("FREQ")), ("Semantic", lambda r: pfam(r).startswith("SEMANTIC")),
        ("Reconstruction", lambda r: pfam(r).startswith("RECON")), ("Hybrid", lambda r: pfam(r).startswith("HYBRID")),
        ("Other", lambda r: pfam(r).split()[0] in ("NOISE", "SOURCE", "GEOMETRY"))]
rl.append("\\newcommand{\\PerFamPlot}{" + stacked_cols(share(None, FAMC), labs, "Share of studies (\\%)", 100, 3) + "}")
for nm, test in FAMC:
    for p in PERIODS: R_("PerFam" + nm + p[0], sum(1 for r in per_rows(p) if test(r)))
# strength by family
STR = [("Very high", "cB!85!black", "VERY HIGH"), ("High", "cB!55", "HIGH"), ("Moderate", "cD", "MODERATE"), ("Low", "cC!45", "LOW")]
famorder = ["HYBRID / MULTI-CUE", "SPATIAL / VISUAL ARTIFACT", "FREQUENCY / SPECTRAL", "SEMANTIC / FOUNDATION-MODEL", "RECONSTRUCTION / INVERSION", "SOURCE / FINGERPRINT FORENSIC", "NOISE / RESIDUAL FORENSIC", "GEOMETRY / PHYSICAL CUES"]
famlab = {"HYBRID / MULTI-CUE": "Hybrid", "SPATIAL / VISUAL ARTIFACT": "Spatial", "FREQUENCY / SPECTRAL": "Frequency", "SEMANTIC / FOUNDATION-MODEL": "Semantic",
          "RECONSTRUCTION / INVERSION": "Reconstruction", "SOURCE / FINGERPRINT FORENSIC": "Source", "NOISE / RESIDUAL FORENSIC": "Noise", "GEOMETRY / PHYSICAL CUES": "Geometry"}
sc = {(f, s): sum(1 for r in AR if pfam(r) == f and r["Generalization_Strength"] == s) for f in famorder for _, _, s in STR}
ytl = ",".join(famlab[f] for f in famorder[::-1])
sbody = "".join("\\addplot[fill=" + fc + ", draw=cA] coordinates {" + " ".join(f"({sc[(f, s)]},{famlab[f]})" for f in famorder[::-1]) + "};" for _, fc, s in STR)
rl.append("\\newcommand{\\StrengthPlot}{\\begin{tikzpicture}\\let\\xspace\\empty\\begin{axis}[xbar stacked, width=\\columnwidth-2.4cm, height=6.2cm, bar width=9pt, symbolic y coords={" + ytl + "}, ytick=data, y dir=normal, xmin=0, xmax=40, xlabel={Studies}, xlabel style={font=\\footnotesize}, tick label style={font=\\footnotesize}, axis lines*=left, enlarge y limits=0.08, legend style={font=\\scriptsize, draw=none, at={(0.98,0.5)}, anchor=east}, legend columns=1]" + sbody + "\\legend{" + ",".join(n for n, _, _ in STR) + "}\\end{axis}\\end{tikzpicture}}")
for f in famorder:
    n = sum(sc[(f, s)] for _, _, s in STR); hv = sc[(f, "VERY HIGH")] + sc[(f, "HIGH")]
    R_("HVn" + famlab[f], hv); rl.append("\\newcommand{\\HVpct" + famlab[f] + "}{" + f"{100*hv/n:.0f}" + "}")
# reproducibility
RP = [("Code available", "Code_Available"), ("Dataset available", "Dataset_Available"), ("Ablation study", "Ablation_Study")]
rpb = ""
for nm, fc, v in [("Yes", "cB!70", "YES"), ("Unclear", "cD", "UNCLEAR"), ("No", "cC!45", "NO")]:
    rpb += "\\addplot[fill=" + fc + ", draw=cA] coordinates {" + " ".join(f"({sum(1 for r in AR if r[h] == v)},{l.replace(' ', '')})" for l, h in RP[::-1]) + "};"
rl.append("\\newcommand{\\ReproPlot}{\\begin{tikzpicture}\\let\\xspace\\empty\\begin{axis}[xbar stacked, width=\\columnwidth, height=3.4cm, bar width=11pt, symbolic y coords={" + ",".join(l.replace(' ', '') for l, _ in RP[::-1]) + "}, yticklabels={" + ",".join(l for l, _ in RP[::-1]) + "}, ytick=data, y dir=normal, xmin=0, xmax=99, xlabel={Studies}, xlabel style={font=\\footnotesize}, tick label style={font=\\footnotesize}, axis lines*=left, enlarge y limits=0.3, legend style={font=\\scriptsize, draw=none, at={(0.5,1.05)}, anchor=south, legend columns=3}]" + rpb + "\\legend{Yes,Unclear,No}\\end{axis}\\end{tikzpicture}}")
R_("NcodeUnclear", sum(1 for r in AR if r["Code_Available"] == "UNCLEAR"))
R_("PctNmethod", pct(len(METH)))
R_("NlevelFrame", sum(1 for r in AR if r["Detection_Level"] == "FRAME")); R_("NlevelImage", sum(1 for r in AR if r["Detection_Level"] == "IMAGE"))
with open(f"{OUT}/results_a_data.tex", "w") as f:
    f.write("% AUTO-GENERATED by scripts/gen_from_workbook.py -- do not edit by hand\n" + "\n".join(rl) + "\n")

# ---------------- Section 6 figures: RQ2 datasets, generators, protocols, metrics ----------------
import statistics as _st
ql = []
def Q_(name, val): ql.append(f"\\newcommand{{\\{name}}}{{{val}\\xspace}}")
def _txt(r, fs): return " ; ".join(str(r[f] or "") for f in fs)
_GEN = [("Stable Diffusion", r"stable diffusion|\bsd ?v?\d|\bsd\b|sdxl|\bsd3"), ("StyleGAN", r"stylegan"), ("DALL-E", r"dall"), ("BigGAN", r"biggan"), ("Midjourney", r"midjourney"),
        ("GLIDE", r"glide"), ("ADM", r"\badm\b"), ("Wukong", r"wukong"), ("StarGAN", r"stargan"), ("ProGAN", r"progan"), ("CycleGAN", r"cyclegan"), ("VQDM", r"vqdm"),
        ("GauGAN", r"gaugan"), ("Face-swap deepfakes", r"deepfake|faceswap|face2face|neuraltextures|fsgan|simswap"), ("Latent diffusion", r"\bldm\b|latent diffusion"),
        ("DDPM/DDIM", r"ddpm|ddim|iddpm"), ("FLUX", r"flux"), ("Imagen/Ideogram/Kandinsky", r"imagen|ideogram|kandinsky|playground"), ("Firefly", r"firefly"), ("DiT/PixArt", r"\bdit\b|pixart")]
_gc = {k: sum(1 for r in AR if _re.search(v, _txt(r, ["Test_Generator_s"]).lower())) for k, v in _GEN}
for k, v in _gc.items(): Q_("Gen" + _re.sub(r"[^A-Za-z]", "", k), v)
ql.append("\\newcommand{\\GenNamedPlot}{" + hbar(sorted(_gc.items(), key=lambda x: -x[1])[:14], "Studies naming the generator in a test set", 62, "6.3cm") + "}")
_nfam = [sum(1 for k, v in _GEN if _re.search(v, _txt(r, ["Test_Generator_s"]).lower())) for r in AR]
Q_("GenFamMedian", int(_st.median(_nfam))); Q_("GenFamZero", sum(1 for x in _nfam if x == 0)); Q_("GenFamOne", sum(1 for x in _nfam if x == 1))
Q_("GenFamTwoThree", sum(1 for x in _nfam if 2 <= x <= 3)); Q_("GenFamFourSeven", sum(1 for x in _nfam if 4 <= x <= 7)); Q_("GenFamEightPlus", sum(1 for x in _nfam if x >= 8))
Q_("GenFamNames", len(_GEN))
_bins = [("0", 0, 0), ("1", 1, 1), ("2-3", 2, 3), ("4-7", 4, 7), ("8+", 8, 99)]
bc = [(l, sum(1 for x in _nfam if lo <= x <= hi)) for l, lo, hi in _bins]
ql.append("\\newcommand{\\GenBreadthPlot}{\\begin{tikzpicture}\\let\\xspace\\empty\\begin{axis}[ybar, width=\\columnwidth, height=4.4cm, bar width=18pt, symbolic x coords={" + ",".join(l for l, _ in bc) + "}, xtick=data, ymin=0, ymax=38, xlabel={Generator families named in the test sets}, ylabel={Studies}, label style={font=\\footnotesize}, tick label style={font=\\footnotesize}, nodes near coords, nodes near coords style={font=\\scriptsize}, axis lines*=left, enlarge x limits=0.15]\\addplot[fill=cB!70, draw=cA] coordinates {" + " ".join(f"({l},{v})" for l, v in bc) + "};\\end{axis}\\end{tikzpicture}}")
_REAL = [("ImageNet", r"imagenet"), ("LSUN", r"lsun"), ("CelebA/CelebA-HQ", r"celeba(?!-?df)"), ("COCO", r"coco"), ("FFHQ", r"ffhq"), ("LAION", r"laion"), ("Flickr/web", r"flickr|web|google|unsplash|pexels"), ("Camera/RAW sets", r"raise|dresden|vision|camera|raw|fivek"), ("FaceForensics", r"faceforensics|ff\+\+"), ("CIFAR", r"cifar")]
_rc = {k: sum(1 for r in AR if _re.search(v, _txt(r, ["Training_Real_Source", "Test_Real_Source_s"]).lower())) for k, v in _REAL}
for k, v in _rc.items(): Q_("Real" + _re.sub(r"[^A-Za-z]", "", k), v)
ql.append("\\newcommand{\\RealSrcPlot}{" + hbar(sorted(_rc.items(), key=lambda x: -x[1]), "Studies using the source for real images", 48, "5.0cm") + "}")
Q_("PctRealImageNet", pct(_rc["ImageNet"]) + "\\%")
_tp = lambda r: _re.search(r"progan", str(r["Training_Generator_s"] or "").lower())
_ts = lambda r: _re.search(r"stable diffusion|\bsd ?v?1|sdxl", str(r["Training_Generator_s"] or "").lower())
_tf_ = lambda r: "training-free" in (str(r["Training_Dataset"] or "") + str(r["Training_Generator_s"] or "")).lower() or str(r["Training_Generator_s"] or "").strip().lower() in ("none", "n/a")
Q_("TrainProGAN", sum(1 for r in AR if _tp(r))); Q_("TrainSD", sum(1 for r in AR if _ts(r))); Q_("TrainBoth", sum(1 for r in AR if _tp(r) and _ts(r)))
Q_("TrainFree", sum(1 for r in AR if _tf_(r))); Q_("TrainOneOfTwo", sum(1 for r in AR if _tp(r) or _ts(r)))
Q_("TrainOther", sum(1 for r in AR if not (_tp(r) or _ts(r) or _tf_(r)))); Q_("PctTrainOneOfTwo", pct(sum(1 for r in AR if _tp(r) or _ts(r))) + "\\%")
_BN = [("GenImage", "GenImage"), ("ForenSynths/ProGAN sets", "ForenSynths"), ("UFD", "UFD"), ("DiffusionForensics", "DiffForensics"), ("FaceForensics", "FFpp"), ("Celeb-DF", "CelebDF"), ("DFDC", "DFDC"), ("CIFAKE", "CIFAKE"), ("Synthbuster", "Synthbuster"), ("Chameleon", "Chameleon"), ("AIGCDetectBenchmark", "AIGCDB")]
ql.append("\\newcommand{\\BenchPlot}{" + hbar([(l, len(BL[k])) for l, k in _BN], "Studies naming the benchmark", 44, "5.8cm") + "}")
_BP = {"GenImage": r"genimage", "ForenSynths/ProGAN sets": r"forensynth|cnn-?spot|wang et al|progan", "UFD": r"universalfake|ojha|ufd", "FaceForensics": r"faceforensics|ff\+\+"}
_series = []
for (nm, rx), (col, mk) in zip(_BP.items(), [("cB!85!black", "square*"), ("cC!80!black", "triangle*"), ("cA", "diamond*"), ("gray", "o")]):
    vals = [round(100 * sum(1 for r in per_rows(p) if _re.search(rx, _txt(r, ["Training_Dataset", "Test_Dataset_s", "Training_Generator_s", "Test_Generator_s"]).lower())) / max(1, len(per_rows(p))), 1) for p in PERIODS]
    for p, r_ in zip(PERIODS, [sum(1 for r in per_rows(p) if _re.search(rx, _txt(r, ["Training_Dataset", "Test_Dataset_s", "Training_Generator_s", "Test_Generator_s"]).lower())) for p in PERIODS]):
        Q_("PerBench" + _re.sub(r"[^A-Za-z]", "", nm) + p[0], r_)
    _series.append((nm, col, mk, vals))
ticks = ",".join("{" + t + "\\\\$n{=}\\PerN" + k + "$}" for k, t, _ in labs)
ql.append("\\newcommand{\\BenchPeriodPlot}{\\begin{tikzpicture}\\let\\xspace\\empty\\begin{axis}[width=\\columnwidth, height=5.2cm, symbolic x coords={a,b,c,d}, xtick=data, xticklabels={" + ticks + "}, xticklabel style={align=center, font=\\scriptsize}, ymin=0, ymax=62, ylabel={Studies naming it (\\%)}, label style={font=\\footnotesize}, tick label style={font=\\footnotesize}, enlarge x limits=0.12, axis lines*=left, grid=major, grid style={dotted,gray!50}, legend style={font=\\scriptsize, draw=none, at={(0.02,0.98)}, anchor=north west}]" +
          "".join(f"\\addplot[color={c}, mark={m}, thick] coordinates {{{' '.join(f'({p[0]},{v})' for p, v in zip(PERIODS, vals))}}};" for _, c, m, vals in _series) + "\\legend{" + ",".join(n for n, _, _, _ in _series) + "}\\end{axis}\\end{tikzpicture}}")
_mt = lambda r: str(r["Metrics"] or "").lower()
_tfm = lambda r: _re.search(r"auc|auroc|\bm?ap\b|average precision|\beer\b", _mt(r))
_acc = lambda r: _re.search(r"\bacc|accuracy", _mt(r))
_prf = lambda r: _re.search(r"\bf1|f-?score|precision|recall", _mt(r))
for nm, fn in [("TfAll", lambda r: _tfm(r) and _acc(r) and _prf(r)), ("TfAcc", lambda r: _tfm(r) and _acc(r) and not _prf(r)), ("TfOnly", lambda r: _tfm(r) and not _acc(r) and not _prf(r)),
               ("TfPrf", lambda r: _tfm(r) and not _acc(r) and _prf(r)), ("AccPrf", lambda r: not _tfm(r) and _acc(r) and _prf(r)), ("AccOnly", lambda r: not _tfm(r) and _acc(r) and not _prf(r)),
               ("NoneOfThese", lambda r: not _tfm(r) and not _acc(r) and not _prf(r))]:
    Q_("MetCat" + nm, sum(1 for r in AR if fn(r)))
for p in PERIODS:
    Q_("PerTf" + p[0], sum(1 for r in per_rows(p) if _tfm(r)))
_mser = [("Reports a threshold-free metric", "cB!70", [round(100 * sum(1 for r in per_rows(p) if _tfm(r)) / len(per_rows(p)), 1) for p in PERIODS]),
         ("No threshold-free metric", "cC!35", [round(100 * sum(1 for r in per_rows(p) if not _tfm(r)) / len(per_rows(p)), 1) for p in PERIODS])]
ql.append("\\newcommand{\\MetricPeriodPlot}{" + stacked_cols(_mser, labs, "Share of studies (\\%)", 100, 1) + "}")
_mu = [("Accuracy", "ACC"), ("F1", "Fone"), ("AUC", "AUC"), ("Average precision", "AP"), ("TPR/FPR", "TPRFPR"), ("EER", "EER")]
ql.append("\\newcommand{\\MetricUsePlot}{" + hbar([(l, stats["Met" + k]) for l, k in _mu], "Studies reporting the metric", 100, "3.6cm") + "}")
_nrf = [("Test sample size", "Test_Sample_Size"), ("Training sample size", "Training_Sample_Size"), ("Training generators", "Training_Generator_s"), ("Test generators", "Test_Generator_s")]
_nr = [(l, sum(1 for r in AR if _re.search(r"\bNR\b", str(r[f] or "")))) for l, f in _nrf]
for (l, v), (_, f) in zip(_nr, _nrf): Q_("NR" + _re.sub(r"[^A-Za-z]", "", l), v)
ql.append("\\newcommand{\\NRPlot}{" + hbar(_nr, "Studies marking the field (partly) not reported", 48, "3.4cm") + "}")
Q_("PctNRTestsamplesize", pct(_nr[0][1]) + "\\%")
Q_("NformatMention", sum(1 for r in AR if _re.search(r"format|jpeg.{0,12}png|png.{0,12}jpeg|compression bias|file[- ]type", _txt(r, ["Main_Limitations", "Threats_to_Validity", "Generalization_Failure_Evidence"]).lower())))
with open(f"{OUT}/results_b_data.tex", "w") as f:
    f.write("% AUTO-GENERATED by scripts/gen_from_workbook.py -- do not edit by hand\n" + "\n".join(ql) + "\n")

# ---------------- Section 7 figures: RQ3 generalization ----------------
cl = []
def G_(name, val): cl.append(f"\\newcommand{{\\{name}}}{{{val}\\xspace}}")
DIMS = [("Unseen_Generator_Test", "D1", "Unseen generator", "One"), ("Cross_Family_Test", "D2", "Cross-family", "Two"), ("Cross_Dataset_Test", "D3", "Cross-dataset", "Three"),
        ("Cross_Domain_Content_Test", "D4", "Cross-domain", "Four"), ("Open_Set_Test", "D5", "Open set", "Five"), ("Continual_Adaptive_Test", "D6", "Continual", "Six")]
strong = lambda r: r["Generalization_Strength"] in ("HIGH", "VERY HIGH")
# strong vs weak evidence per dimension
sbody = ""
for fc, nm, key in [("cB!70", "Rated high or very high", True), ("cC!35", "Rated moderate or low", False)]:
    sbody += "\\addplot[fill=" + fc + ", draw=cA] coordinates {" + " ".join(f"({sum(1 for r in AR if Y(r, h) and strong(r) == key)},{d} {l})" for h, d, l, _ in DIMS[::-1]) + "};"
for h, d, l, nm in DIMS:
    s = sum(1 for r in AR if Y(r, h) and strong(r)); w = sum(1 for r in AR if Y(r, h) and not strong(r))
    G_("Strong" + nm, s); G_("Weak" + nm, w); cl.append("\\newcommand{\\PctStrong" + nm + "}{" + f"{100*s/(s+w):.0f}" + "}")
cl.append("\\newcommand{\\DimPlot}{\\begin{tikzpicture}\\let\\xspace\\empty\\begin{axis}[xbar stacked, width=\\columnwidth-2.9cm, height=5.2cm, bar width=10pt, symbolic y coords={" + ",".join(f"{d} {l}" for _, d, l, _ in DIMS[::-1]) +
          "}, ytick=data, y dir=normal, xmin=0, xmax=80, xlabel={Studies testing the dimension}, xlabel style={font=\\footnotesize}, tick label style={font=\\footnotesize}, axis lines*=left, enlarge y limits=0.1, legend style={font=\\scriptsize, draw=none, at={(0.5,1.03)}, anchor=south, legend columns=2}]" + sbody + "\\legend{Rated high or very high,Rated moderate or low}\\end{axis}\\end{tikzpicture}}")
# combinations (UpSet-style)
combo = collections.Counter(tuple(i for i, (h, *_ ) in enumerate(DIMS) if Y(r, h)) for r in AR)
top = sorted(combo.items(), key=lambda kv: (-kv[1], kv[0]))[:10]
G_("ComboTopN", sum(v for _, v in top)); G_("ComboNone", combo.get((), 0)); G_("ComboOneToFour", combo.get((0, 1, 2, 3), 0)); G_("ComboOneToFive", combo.get((0, 1, 2, 3, 4), 0)); G_("ComboNdistinct", len(combo))
tk = ["\\begin{tikzpicture}[font=\\scriptsize]"]
for r_, (_, d, l, _) in enumerate(DIMS):
    tk.append(f"\\node[anchor=east] at (-0.05,{-0.55 - r_*0.36:.2f}) {{{d}}};")
for i, (cb, n) in enumerate(top):
    x = i * 0.7 + 0.35
    tk.append(f"\\fill[cB!70, draw=cA] ({x-0.22:.2f},0) rectangle ({x+0.22:.2f},{n*0.1:.2f}); \\node[above] at ({x:.2f},{n*0.1:.2f}) {{{n}}};")
    for r_ in range(6):
        y = -0.55 - r_*0.36
        tk.append(f"\\fill[{'cA' if r_ in cb else 'cD'}] ({x:.2f},{y:.2f}) circle (2.6pt);")
    if cb:
        tk.append(f"\\draw[cA, thick] ({x:.2f},{-0.55 - min(cb)*0.36:.2f}) -- ({x:.2f},{-0.55 - max(cb)*0.36:.2f});")
tk.append("\\draw[cA] (0,0) -- (" + f"{len(top)*0.7:.2f}" + ",0);\\end{tikzpicture}")
cl.append("\\newcommand{\\ComboPlot}{" + "".join(tk) + "}")
# ladder reach by period
def ladder(r):
    if not Y(r, "Unseen_Generator_Test"): return 0
    if not (Y(r, "Cross_Family_Test") and Y(r, "Cross_Dataset_Test")): return 1
    if r["Robustness_Tested"] != "YES": return 2
    if not Y(r, "Open_Set_Test"): return 3
    return 5 if Y(r, "Continual_Adaptive_Test") else 4
_ls = []
for k, (nm, col, mk) in enumerate([("L1 unseen generator", "cB!85!black", "square*"), ("L2 + family and dataset", "cC!80!black", "triangle*"), ("L3 + robustness", "cA", "diamond*"), ("L4 + open set", "gray", "o")], start=1):
    vals = []
    for p in PERIODS:
        n_ = sum(1 for r in per_rows(p) if ladder(r) >= k); G_("LadPerL" + ["one", "two", "three", "four"][k - 1] + p[0], n_)
        vals.append(round(100 * n_ / len(per_rows(p)), 1))
    _ls.append((nm, col, mk, vals))
cl.append("\\newcommand{\\LadderPeriodPlot}{\\begin{tikzpicture}\\let\\xspace\\empty\\begin{axis}[width=\\columnwidth, height=5.2cm, symbolic x coords={a,b,c,d}, xtick=data, xticklabels={" + ticks + "}, xticklabel style={align=center, font=\\scriptsize}, ymin=0, ymax=100, ylabel={Studies reaching the level (\\%)}, label style={font=\\footnotesize}, tick label style={font=\\footnotesize}, enlarge x limits=0.12, axis lines*=left, grid=major, grid style={dotted,gray!50}, legend style={font=\\scriptsize, draw=none, at={(0.5,1.03)}, anchor=south, legend columns=2}]" +
          "".join(f"\\addplot[color={c}, mark={m}, thick] coordinates {{{' '.join(f'({p[0]},{v})' for p, v in zip(PERIODS, vals))}}};" for _, c, m, vals in _ls) + "\\legend{" + ",".join(n for n, _, _, _ in _ls) + "}\\end{axis}\\end{tikzpicture}}")
# breadth x strength
brd = lambda r: sum(1 for h, *_ in DIMS if Y(r, h))
rows_ = []
for b in range(5):
    rs = [r for r in AR if min(brd(r), 4) == b]
    cts = [sum(1 for r in rs if r["Generalization_Strength"] == s) for s in ("VERY HIGH", "HIGH", "MODERATE", "LOW")]
    rows_.append(("4 or more" if b == 4 else str(b)) + " & " + " & ".join(str(c) for c in cts) + f" & {len(rs)} \\\\")
cl.append("\\newcommand{\\BreadthRows}{" + "\n".join(rows_) + "}")
b4 = [r for r in AR if brd(r) >= 4]
G_("BreadthFourN", len(b4)); G_("BreadthFourStrong", sum(1 for r in b4 if strong(r))); G_("BreadthFourModerate", sum(1 for r in b4 if r["Generalization_Strength"] == "MODERATE")); G_("BreadthFourLow", sum(1 for r in b4 if r["Generalization_Strength"] == "LOW"))
# family x dimension heat map
FD = [("Hybrid", lambda r: pfam(r).startswith("HYBRID")), ("Spatial", lambda r: pfam(r).startswith("SPATIAL")), ("Frequency", lambda r: pfam(r).startswith("FREQ")),
      ("Semantic", lambda r: pfam(r).startswith("SEMANTIC")), ("Reconstruction", lambda r: pfam(r).startswith("RECON")), ("Other", lambda r: pfam(r).split()[0] in ("NOISE", "SOURCE", "GEOMETRY"))]
cells = []
for fi, (fn, ft) in enumerate(FD):
    rs = [r for r in AR if ft(r)]
    G_("DFn" + fn, len(rs))
    for ci, (h, d, l, nm) in enumerate(DIMS):
        n_ = sum(1 for r in rs if Y(r, h)); G_("DF" + fn + nm, n_); cells.append(f"\\dimcell{{{fi}}}{{{ci}}}{{{round(100*n_/len(rs))}}}{{{n_}}}")
    n0 = sum(1 for r in rs if brd(r) == 0); G_("DF" + fn + "None", n0); cells.append(f"\\dimcell{{{fi}}}{{6}}{{{round(100*n0/len(rs))}}}{{{n0}}}")
cl.append("\\newcommand{\\DimFamCells}{" + "".join(cells) + "}")
# claim audit
clm = lambda r: r["Authors_Claimed_Generalization"] == "YES"; sup = lambda r: r["Evidence_Supports_Generalization"] == "YES"
G_("ClaimSupp", sum(1 for r in AR if clm(r) and sup(r))); G_("ClaimUnsupp", sum(1 for r in AR if clm(r) and not sup(r))); G_("NoClaim", sum(1 for r in AR if not clm(r)))
mm = [r for r in AR if clm(r) and not sup(r)]
G_("MismNoTest", sum(1 for r in mm if r["Generalization_Protocol"] == "NONE")); G_("MismTested", sum(1 for r in mm if r["Generalization_Protocol"] != "NONE"))
G_("MismLow", sum(1 for r in mm if r["Generalization_Strength"] == "LOW")); G_("MismModerate", sum(1 for r in mm if r["Generalization_Strength"] == "MODERATE"))
G_("MismHybrid", sum(1 for r in mm if pfam(r).startswith("HYBRID"))); G_("MismSpatial", sum(1 for r in mm if pfam(r).startswith("SPATIAL")))
G_("PctClaim", pct(stats["Nclaim"]))
for p in PERIODS:
    G_("ClaimPer" + p[0], sum(1 for r in per_rows(p) if clm(r))); G_("MismPer" + p[0], sum(1 for r in per_rows(p) if clm(r) and not sup(r)))
_cs = [("Claim supported by own experiments", "cB!70", [round(100 * sum(1 for r in per_rows(p) if clm(r) and sup(r)) / len(per_rows(p)), 1) for p in PERIODS]),
       ("Claim without supporting experiment", "cC!55", [round(100 * sum(1 for r in per_rows(p) if clm(r) and not sup(r)) / len(per_rows(p)), 1) for p in PERIODS]),
       ("No generalization claim", "cD", [round(100 * sum(1 for r in per_rows(p) if not clm(r)) / len(per_rows(p)), 1) for p in PERIODS])]
cl.append("\\newcommand{\\ClaimPeriodPlot}{" + stacked_cols(_cs, labs, "Share of studies (\\%)", 100, 1) + "}")
G_("NstrongTotal", sum(1 for r in AR if strong(r)))
with open(f"{OUT}/results_c_data.tex", "w") as f:
    f.write("% AUTO-GENERATED by scripts/gen_from_workbook.py -- do not edit by hand\n" + "\n".join(cl) + "\n")

# ---------------- Section 8 figures: RQ4 robustness ----------------
dl = []
def D_(name, val): dl.append(f"\\newcommand{{\\{name}}}{{{val}\\xspace}}")
RFL = [("JPEG_Compression", "JPEG"), ("Blur", "Blur"), ("Resize", "Resize"), ("Noise", "Noise"), ("Social_Media_Postprocessing", "Social network"), ("Adversarial_Attack", "Adversarial")]
rob_any = lambda r: r["Robustness_Tested"] == "YES"
cnt = [(l, sum(1 for r in AR if Y(r, h))) for h, l in RFL] + [("Any robustness test", sum(1 for r in AR if rob_any(r)))]
dl.append("\\newcommand{\\RobCondPlot}{" + hbar(cnt, "Studies testing the condition", 66, "4.2cm") + "}")
rs_ = []
for nm, col, mk, fn in [("Any robustness test", "cB!85!black", "square*", rob_any), ("JPEG", "cC!80!black", "triangle*", lambda r: Y(r, "JPEG_Compression")),
                        ("Social network", "cA", "diamond*", lambda r: Y(r, "Social_Media_Postprocessing")), ("Adversarial", "gray", "o", lambda r: Y(r, "Adversarial_Attack"))]:
    vals = []
    for p in PERIODS:
        n_ = sum(1 for r in per_rows(p) if fn(r)); D_("RobPer" + _re.sub(r"[^A-Za-z]", "", nm) + p[0], n_); vals.append(round(100 * n_ / len(per_rows(p)), 1))
    rs_.append((nm, col, mk, vals))
dl.append("\\newcommand{\\RobPeriodPlot}{\\begin{tikzpicture}\\let\\xspace\\empty\\begin{axis}[width=\\columnwidth, height=5.2cm, symbolic x coords={a,b,c,d}, xtick=data, xticklabels={" + ticks + "}, xticklabel style={align=center, font=\\scriptsize}, ymin=0, ymax=100, ylabel={Studies testing it (\\%)}, label style={font=\\footnotesize}, tick label style={font=\\footnotesize}, enlarge x limits=0.12, axis lines*=left, grid=major, grid style={dotted,gray!50}, legend style={font=\\scriptsize, draw=none, at={(0.5,1.03)}, anchor=south, legend columns=2}]" +
          "".join(f"\\addplot[color={c}, mark={m}, thick] coordinates {{{' '.join(f'({p[0]},{v})' for p, v in zip(PERIODS, vals))}}};" for _, c, m, vals in rs_) + "\\legend{" + ",".join(n for n, _, _, _ in rs_) + "}\\end{axis}\\end{tikzpicture}}")
nc = collections.Counter(sum(1 for h, _ in RFL if Y(r, h)) for r in AR)
for k in range(7): D_("RobCount" + ["Zero", "One", "Two", "Three", "Four", "Five", "Six"][k], nc.get(k, 0))
D_("RobOtherOnly", sum(1 for r in AR if rob_any(r) and sum(1 for h, _ in RFL if Y(r, h)) == 0)); D_("RobNone", sum(1 for r in AR if not rob_any(r)))
dl.append("\\newcommand{\\RobCountPlot}{\\begin{tikzpicture}\\let\\xspace\\empty\\begin{axis}[ybar, width=\\columnwidth, height=4.2cm, bar width=14pt, symbolic x coords={0,1,2,3,4,5,6}, xtick=data, ymin=0, ymax=52, xlabel={Coded robustness conditions tested (of six)}, ylabel={Studies}, label style={font=\\footnotesize}, tick label style={font=\\footnotesize}, nodes near coords, nodes near coords style={font=\\scriptsize}, axis lines*=left, enlarge x limits=0.1]\\addplot[fill=cB!70, draw=cA] coordinates {" + " ".join(f"({k},{nc.get(k,0)})" for k in range(7)) + "};\\end{axis}\\end{tikzpicture}}")
gen_t = lambda r: r["Generalization_Protocol"] != "NONE"
D_("RobBoth", sum(1 for r in AR if gen_t(r) and rob_any(r))); D_("RobGenOnly", sum(1 for r in AR if gen_t(r) and not rob_any(r)))
D_("RobRobOnly", sum(1 for r in AR if not gen_t(r) and rob_any(r))); D_("RobNeither", sum(1 for r in AR if not gen_t(r) and not rob_any(r)))
for fn_, ft in FD:
    rs = [r for r in AR if ft(r)]
    D_("RobFam" + fn_, sum(1 for r in rs if rob_any(r)))
D_("RobBothOsnAdv", sum(1 for r in AR if Y(r, "Social_Media_Postprocessing") and Y(r, "Adversarial_Attack")))
# JPEG: clean vs compressed accuracy within a study
jp_pairs = []
for r in RR:
    if r["Setting"] == "robustness" and r["Shift"] == "jpeg" and r["Metric"] == "ACC":
        idv = [x["Value_pct"] for x in RR if x["Record_ID"] == r["Record_ID"] and x["Setting"] == "in-domain" and x["Metric"] == "ACC"]
        if idv: jp_pairs.append((r["Record_ID"], r["Target"], idv[0], r["Value_pct"]))
jp_pairs.sort(key=lambda x: -(x[2] - x[3]))
D_("NjpegPairs", len(jp_pairs)); dr = [x[2] - x[3] for x in jp_pairs]
D_("MedJpegDrop", f"{_st.median(dr):.1f}"); D_("MinJpegDrop", f"{min(dr):.1f}"); D_("MaxJpegDrop", f"{max(dr):.1f}")
D_("NdropTenJpeg", sum(1 for d in dr if d >= 10)); D_("MinJpegVal", f"{min(x[3] for x in jp_pairs):.1f}")
D_("NjpegAll", sum(1 for r in RR if r["Setting"] == "robustness" and r["Shift"] == "jpeg"))
tkj = ["\\begin{tikzpicture}[font=\\scriptsize]"]
for i, (rid, tg, a, b) in enumerate(jp_pairs):
    y = -i * 0.52; q = _re.search(r"Q[F=]? ?(\d+)", tg); lab = f"\\cite{{{rid}}}" + (f" Q{q.group(1)}" if q else " (JPEG)")
    xa, xb = (a - 60) * 0.145, (b - 60) * 0.145
    tkj.append(f"\\node[anchor=east] at (-0.1,{y:.2f}) {{{lab}}}; \\draw[cA, thick] ({xa:.2f},{y:.2f}) -- ({xb:.2f},{y:.2f}); \\fill[cB!80!black] ({xa:.2f},{y:.2f}) circle (2.6pt); \\fill[cC!85!black] ({xb:.2f},{y:.2f}) circle (2.6pt); \\node[anchor=west, text=cC!85!black] at ({max(xa, xb)+0.12:.2f},{y:.2f}) {{-{a-b:.1f}}};")
yb = -len(jp_pairs) * 0.52 + 0.1
tkj.append(f"\\draw[cA] (0,{yb:.2f}) -- (5.8,{yb:.2f}); \\foreach \\v in {{60,70,80,90,100}} {{\\draw[cA] ({{(\\v-60)*0.145}},{yb:.2f}) -- ++(0,-0.08); \\node[below] at ({{(\\v-60)*0.145}},{yb-0.08:.2f}) {{\\v}};}}")
tkj.append(f"\\node at (2.9,{yb-0.62:.2f}) {{Accuracy (\\%)}}; \\fill[cB!80!black] (0.3,{yb-1.05:.2f}) circle (2.6pt); \\node[anchor=west] at (0.45,{yb-1.05:.2f}) {{in-domain}}; \\fill[cC!85!black] (2.1,{yb-1.05:.2f}) circle (2.6pt); \\node[anchor=west] at (2.25,{yb-1.05:.2f}) {{after JPEG (points lost)}};\\end{{tikzpicture}}")
dl.append("\\newcommand{\\JpegPairPlot}{" + "".join(tkj) + "}")
with open(f"{OUT}/results_d_data.tex", "w") as f:
    f.write("% AUTO-GENERATED by scripts/gen_from_workbook.py -- do not edit by hand\n" + "\n".join(dl) + "\n")

# ---------------- Section 9 figures: reported performance under shift ----------------
import math as _math
el = []
def E_(name, val): el.append(f"\\newcommand{{\\{name}}}{{{val}\\xspace}}")
def sign_p(pos, n):
    x = max(pos, n - pos); return min(1.0, 2 * sum(_math.comb(n, k) for k in range(x, n + 1)) / 2 ** n)
vals_all = lambda rid, st_, ms: [r["Value_pct"] for r in RR if r["Record_ID"] == rid and r["Setting"] == st_ and r["Metric"] in ms]
PA = {k: [] for k in ("Acc", "Rank")}
for rid in sorted({r["Record_ID"] for r in RR if r["Pair_Eligible"] == "Y"}):
    i, o = vals_all(rid, "in-domain", ("ACC",)), vals_all(rid, "ood", ("ACC",))
    if i and o: PA["Acc"].append(dict(id=rid, i=sum(i)/len(i), o=sum(o)/len(o), omin=min(o), no=len(o)))
    for m in ("AP", "AUC"):
        i, o = vals_all(rid, "in-domain", (m,)), vals_all(rid, "ood", (m,))
        if i and o: PA["Rank"].append(dict(id=rid, i=sum(i)/len(i), o=sum(o)/len(o), omin=min(o), no=len(o))); break
for k, P in PA.items():
    d = [x["i"] - x["o"] for x in P]
    E_("PosDrop" + k, sum(1 for x in d if x > 0)); E_("MedDropWorst" + k, f"{_st.median([x['i'] - x['omin'] for x in P]):.1f}")
    E_("NsingleOod" + k, sum(1 for x in P if x["no"] == 1))
ACCd = {x["id"]: x["i"] - x["o"] for x in PA["Acc"]}; RNKd = {x["id"]: x["i"] - x["o"] for x in PA["Rank"]}
both = sorted(set(ACCd) & set(RNKd)); gt = sum(1 for r in both if ACCd[r] > RNKd[r])
E_("NbothPairs", len(both)); E_("NbothAccGreater", gt); E_("PsignBoth", f"{sign_p(gt, len(both)):.3f}")
E_("NthrSame", len(thr)); E_("PsignThr", f"{sign_p(sum(1 for _, v in thr if v > 0), len(thr)):.4f}")
streng = {r["Record_ID"]: r["Generalization_Strength"] for r in AR}
hv = [ACCd[r] for r in ACCd if streng[r] in ("HIGH", "VERY HIGH")]; ml = [ACCd[r] for r in ACCd if streng[r] in ("MODERATE", "LOW")]
E_("NaccHV", len(hv)); E_("MedDropAccHV", f"{_st.median(hv):.1f}"); E_("NaccML", len(ml)); E_("MedDropAccML", f"{_st.median(ml):.1f}")
for fam_, ft_ in [("Hybrid", lambda r: pfam(next(a for a in AR if a["Record_ID"] == r)).startswith("HYBRID")), ("Spatial", lambda r: pfam(next(a for a in AR if a["Record_ID"] == r)).startswith("SPATIAL"))]:
    x = [ACCd[r] for r in ACCd if ft_(r)]; E_("NaccFam" + fam_, len(x)); E_("MedDropAccFam" + fam_, f"{_st.median(x):.1f}")
idt = [r["Value_pct"] for r in RR if r["Setting"] == "in-domain" and r["Metric"] == "ACC" and next(a for a in AR if a["Record_ID"] == r["Record_ID"])["Generalization_Protocol"] != "NONE"]
E_("NidTransferVals", len(idt)); E_("MedIDtransfer", f"{_st.median(idt):.1f}")
# histogram of all out-of-distribution accuracy values
_oo = [r["Value_pct"] for r in RR if r["Setting"] == "ood" and r["Metric"] == "ACC"]
_bn = [(40, 50), (50, 60), (60, 70), (70, 80), (80, 90), (90, 100.01)]
_bc = [sum(1 for v in _oo if lo <= v < hi) for lo, hi in _bn]
_lb = ["40-50", "50-60", "60-70", "70-80", "80-90", "90-100"]
el.append("\\newcommand{\\OodHistPlot}{\\begin{tikzpicture}\\let\\xspace\\empty\\begin{axis}[ybar, width=\\columnwidth, height=4.4cm, bar width=18pt, symbolic x coords={" + ",".join(_lb) + "}, xtick=data, ymin=0, ymax=34, xlabel={Out-of-distribution accuracy (\\%)}, ylabel={Reported values}, label style={font=\\footnotesize}, tick label style={font=\\footnotesize}, nodes near coords, nodes near coords style={font=\\scriptsize}, axis lines*=left, enlarge x limits=0.1]\\addplot[fill=cC!45, draw=cA] coordinates {" + " ".join(f"({l},{c})" for l, c in list(zip(_lb, _bc))[:3]) + "};\\addplot[fill=cB!70, draw=cA] coordinates {" + " ".join(f"({l},{c})" for l, c in list(zip(_lb, _bc))[3:]) + "};\\end{axis}\\end{tikzpicture}}")
# dumbbell plots
def dumbbell(P, lo, hi, sc, label):
    P = sorted(P, key=lambda x: -(x["i"] - x["o"])); tk = ["\\begin{tikzpicture}[font=\\scriptsize]"]
    for k, x in enumerate(P):
        y = -k * 0.27; xa, xb = (x["i"] - lo) * sc, (x["o"] - lo) * sc
        tk.append(f"\\node[anchor=east] at (-0.08,{y:.2f}) {{\\cite{{{x['id']}}}}}; \\draw[cA, thick] ({xa:.2f},{y:.2f}) -- ({xb:.2f},{y:.2f}); \\fill[cB!80!black] ({xa:.2f},{y:.2f}) circle (2.2pt); \\fill[cC!85!black] ({xb:.2f},{y:.2f}) circle (2.2pt); \\node[anchor=west, text=cC!85!black] at ({max(xa, xb)+0.1:.2f},{y:.2f}) {{-{x['i']-x['o']:.1f}}};")
    yb = -len(P) * 0.27 + 0.1; ax = (hi - lo) * sc
    tk.append(f"\\draw[cA] (0,{yb:.2f}) -- ({ax:.2f},{yb:.2f}); \\foreach \\v in {{{','.join(str(v) for v in range(lo, hi + 1, 10))}}} {{\\draw[cA] ({{(\\v-{lo})*{sc}}},{yb:.2f}) -- ++(0,-0.07); \\node[below] at ({{(\\v-{lo})*{sc}}},{yb-0.07:.2f}) {{\\v}};}} \\node at ({ax/2:.2f},{yb-0.62:.2f}) {{{label}}};\\end{{tikzpicture}}")
    return "".join(tk)
el.append("\\newcommand{\\PairAccPlot}{" + dumbbell(PA["Acc"], 50, 100, 0.125, "Accuracy (\\%)") + "}")
el.append("\\newcommand{\\PairRankPlot}{" + dumbbell(PA["Rank"], 60, 100, 0.155, "AP or AUC (\\%)") + "}")
# strip plot of drops with IQR and median
def strip():
    tk = ["\\begin{tikzpicture}[font=\\scriptsize]"]; sc = 0.15
    for row, (k, nm) in enumerate([("Acc", "Accuracy"), ("Rank", "AP / AUC")]):
        d = sorted(x["i"] - x["o"] for x in PA[k]); y0 = -row * 1.5
        q1, q3, md = statistics.quantiles(d, n=4)[0], statistics.quantiles(d, n=4)[2], _st.median(d)
        tk.append(f"\\node[anchor=east] at (-0.1,{y0:.2f}) {{{nm}}}; \\fill[cD!70] ({q1*sc:.2f},{y0-0.45:.2f}) rectangle ({q3*sc:.2f},{y0+0.45:.2f}); \\draw[cA, very thick] ({md*sc:.2f},{y0-0.45:.2f}) -- ({md*sc:.2f},{y0+0.45:.2f});")
        for j, v in enumerate(d): tk.append(f"\\fill[{'cC!85!black' if k == 'Acc' else 'cB!80!black'}] ({v*sc:.2f},{y0 + (j % 3 - 1)*0.2:.2f}) circle (2pt);")
    yb = -1.5 - 0.7; tk.append(f"\\draw[cA] (0,{yb:.2f}) -- ({45*sc:.2f},{yb:.2f}); \\foreach \\v in {{0,10,20,30,40}} {{\\draw[cA] ({{\\v*{sc}}},{yb:.2f}) -- ++(0,-0.07); \\node[below] at ({{\\v*{sc}}},{yb-0.07:.2f}) {{\\v}};}} \\node at ({45*sc/2:.2f},{yb-0.65:.2f}) {{Within-study drop, in-domain minus out-of-distribution (points)}};\\end{{tikzpicture}}")
    return "".join(tk)
el.append("\\newcommand{\\DropStripPlot}{" + strip() + "}")
# slope chart for studies with both pairs
def slope():
    tk = ["\\begin{tikzpicture}[font=\\scriptsize]"]; sc = 0.13
    for r in both:
        a, b = RNKd[r] * sc, ACCd[r] * sc; col = "cC!70" if ACCd[r] > RNKd[r] else "gray"
        tk.append(f"\\draw[{col}, thick] (0,{a:.2f}) -- (3,{b:.2f}); \\fill[cB!80!black] (0,{a:.2f}) circle (2pt); \\fill[cC!85!black] (3,{b:.2f}) circle (2pt);")
    tk.append(f"\\node[above] at (0,{31*sc:.2f}) {{AP / AUC drop}}; \\node[above] at (3,{31*sc:.2f}) {{Accuracy drop}}; \\draw[cA] (-0.4,0) -- (-0.4,{30*sc:.2f}); \\foreach \\v in {{0,10,20,30}} {{\\draw[cA] (-0.4,{{\\v*{sc}}}) -- ++(-0.07,0); \\node[left] at (-0.47,{{\\v*{sc}}}) {{\\v}};}}\\end{{tikzpicture}}")
    return "".join(tk)
el.append("\\newcommand{\\SlopePlot}{" + slope() + "}")
el.append("\\newcommand{\\BothCites}{\\cite{" + ",".join(both) + "}}")
# threshold gap per study
def thrgap():
    P = sorted(thr, key=lambda x: -x[1]); tk = ["\\begin{tikzpicture}[font=\\scriptsize]"]; sc = 0.27; x0 = 1.0
    for k, (rid, v) in enumerate(P):
        y = -k * 0.3; tk.append(f"\\node[anchor=east] at (0.55,{y:.2f}) {{\\cite{{{rid}}}}}; \\fill[{'cB!70' if v > 0 else 'cC!55'}, draw=cA] ({x0:.2f},{y-0.1:.2f}) rectangle ({x0 + v*sc:.2f},{y+0.1:.2f}); \\node[anchor={'west' if v > 0 else 'east'}] at ({x0 + v*sc + (0.06 if v > 0 else -0.06):.2f},{y:.2f}) {{{v:+.1f}}};")
    yb = -len(P) * 0.3 + 0.1; md = _st.median([v for _, v in P])
    tk.append(f"\\draw[cA] ({x0:.2f},0.2) -- ({x0:.2f},{yb:.2f}); \\draw[cC!85!black, dashed] ({x0 + md*sc:.2f},0.2) -- ({x0 + md*sc:.2f},{yb:.2f}); \\node[anchor=south, text=cC!85!black] at ({x0 + md*sc:.2f},0.2) {{median {md:.1f}}}; \\node at ({x0 + 2.2:.2f},{yb-0.4:.2f}) {{AP minus accuracy on identical targets (points)}};\\end{{tikzpicture}}")
    return "".join(tk)
el.append("\\newcommand{\\ThrGapPlot}{" + thrgap() + "}")
with open(f"{OUT}/results_e_data.tex", "w") as f:
    f.write("% AUTO-GENERATED by scripts/gen_from_workbook.py -- do not edit by hand\n" + "\n".join(el) + "\n")

# ---------------- Section 10 figures: failure modes and quality appraisal ----------------
fl_ = []
def F_(name, val): fl_.append(f"\\newcommand{{\\{name}}}{{{val}\\xspace}}")
QAK = [("QA1", "QA1 Task clarity"), ("QA2", "QA2 Dataset transparency"), ("QA3", "QA3 Protocol clarity"), ("QA4", "QA4 Baselines"), ("QA5", "QA5 Generalization"), ("QA6", "QA6 Robustness"), ("QA7", "QA7 Metrics"), ("QA8", "QA8 Limitations")]
QAR = {}
for r in QAws.iter_rows(min_row=12, values_only=True):
    if r[0] and str(r[0]).startswith("R0"): QAR[r[0]] = [float(x) for x in r[2:10]]
qtot = {k: sum(v) for k, v in QAR.items()}; qcore = {k: sum(v[i] for i in (0, 1, 2, 3, 6, 7)) for k, v in QAR.items()}
# per-question stacked bars (1 / 0.5 / 0)
qb = ""
for nm, fc, val in [("Score 1", "cB!70", 1.0), ("Score 0.5", "cD", 0.5), ("Score 0", "cC!45", 0.0)]:
    qb += "\\addplot[fill=" + fc + ", draw=cA] coordinates {" + " ".join(f"({sum(1 for v in QAR.values() if v[i] == val)},{QAK[i][0]})" for i in range(7, -1, -1)) + "};"
fl_.append("\\newcommand{\\QAQuestionPlot}{\\begin{tikzpicture}\\let\\xspace\\empty\\begin{axis}[xbar stacked, width=\\columnwidth-1.2cm, height=5.6cm, bar width=9pt, symbolic y coords={" + ",".join(QAK[i][0] for i in range(7, -1, -1)) + "}, ytick=data, y dir=normal, xmin=0, xmax=99, xlabel={Studies}, xlabel style={font=\\footnotesize}, tick label style={font=\\footnotesize}, axis lines*=left, enlarge y limits=0.08, legend style={font=\\scriptsize, draw=none, at={(0.5,1.03)}, anchor=south, legend columns=3}]" + qb + "\\legend{Score 1,Score 0.5,Score 0}\\end{axis}\\end{tikzpicture}}")
ht = collections.Counter(qtot.values()); bins = [3.0, 3.5, 4.0, 4.5, 5.0, 5.5, 6.0, 6.5, 7.0, 7.5]
fl_.append("\\newcommand{\\QATotalPlot}{\\begin{tikzpicture}\\let\\xspace\\empty\\begin{axis}[ybar, width=\\columnwidth, height=4.4cm, bar width=11pt, symbolic x coords={" + ",".join(str(b) for b in bins) + "}, xtick=data, ymin=0, ymax=30, xlabel={Total appraisal score (maximum 8)}, ylabel={Studies}, label style={font=\\footnotesize}, tick label style={font=\\footnotesize}, nodes near coords, nodes near coords style={font=\\scriptsize}, axis lines*=left, enlarge x limits=0.07]\\addplot[fill=cB!70, draw=cA] coordinates {" + " ".join(f"({b},{ht.get(b, 0)})" for b in bins) + "};\\end{axis}\\end{tikzpicture}}")
F_("QAMax", f"{max(qtot.values()):.1f}"); F_("QAMinScore", f"{min(qtot.values()):.1f}")
F_("QAAtMax", sum(1 for v in qtot.values() if v == max(qtot.values())))
for p in PERIODS: F_("QAMean" + p[0], f"{_st.mean([qtot[r['Record_ID']] for r in per_rows(p)]):.2f}")
for fn_, ft in FD:
    x = [qtot[r["Record_ID"]] for r in AR if ft(r)]; F_("QAFamN" + fn_, len(x)); F_("QAFamMean" + fn_, f"{_st.mean(x):.2f}")
rkk = {"LOW": 1, "MODERATE": 2, "HIGH": 3, "VERY HIGH": 4}
def _rho(x, y):
    def rk(v):
        s = sorted(range(len(v)), key=lambda i: v[i]); rr = [0] * len(v); i = 0
        while i < len(s):
            j = i
            while j + 1 < len(s) and v[s[j + 1]] == v[s[i]]: j += 1
            for k in range(i, j + 1): rr[s[k]] = (i + j) / 2 + 1
            i = j + 1
        return rr
    a, b = rk(x), rk(y); ma, mb = _st.mean(a), _st.mean(b)
    return sum((p - ma) * (q - mb) for p, q in zip(a, b)) / (sum((p - ma) ** 2 for p in a) * sum((q - mb) ** 2 for q in b)) ** 0.5
ids_ = [r["Record_ID"] for r in AR]; sv = [rkk[r["Generalization_Strength"]] for r in AR]
F_("RhoQATotal", f"{_rho([qtot[i] for i in ids_], sv):.2f}"); F_("RhoQACore", f"{_rho([qcore[i] for i in ids_], sv):.2f}")
for s, nm in [("VERY HIGH", "VH"), ("HIGH", "H"), ("MODERATE", "M"), ("LOW", "L")]:
    ii = [r["Record_ID"] for r in AR if r["Generalization_Strength"] == s]
    F_("QAStrN" + nm, len(ii)); F_("QAStrTot" + nm, f"{_st.median([qtot[i] for i in ii]):.1f}"); F_("QAStrCore" + nm, f"{_st.median([qcore[i] for i in ii]):.1f}")
full = [k for k, b in {r: str(QA[r]["QAbasis"]) for r in QA}.items() if "full text" in b]; prox = [k for k in QAR if k not in full]
F_("QAEightFullOne", sum(1 for k in full if QAR[k][7] == 1.0)); F_("QAEightFullHalf", sum(1 for k in full if QAR[k][7] == 0.5)); F_("QAEightFullZero", sum(1 for k in full if QAR[k][7] == 0.0))
F_("QAEightProxHalf", sum(1 for k in prox if QAR[k][7] == 0.5)); F_("QAEightProxN", len(prox))
lo = sorted(qtot, key=lambda k: qtot[k]); hi = sorted(qtot, key=lambda k: -qtot[k])
F_("QALowCites", "\\cite{" + ",".join(k for k in lo if qtot[k] <= 3.0) + "}"); F_("QAHighCites", "\\cite{" + ",".join(k for k in hi if qtot[k] == max(qtot.values())) + "}")
F_("QAMeanOne", f"{_st.mean([v[0] for v in QAR.values()]):.2f}")
with open(f"{OUT}/results_f_data.tex", "w") as f:
    f.write("% AUTO-GENERATED by scripts/gen_from_workbook.py -- do not edit by hand\n" + "\n".join(fl_) + "\n")

# ---------------- Section 11 figures: Discussion ----------------
gl = []
def H_(name, val): gl.append(f"\\newcommand{{\\{name}}}{{{val}\\xspace}}")
def _fac(r):
    return [Y(r, "Unseen_Generator_Test") or Y(r, "Cross_Family_Test"), Y(r, "Cross_Dataset_Test"), Y(r, "Cross_Domain_Content_Test"),
            any(Y(r, f) for f in ("JPEG_Compression", "Blur", "Resize", "Noise", "Social_Media_Postprocessing")), Y(r, "Adversarial_Attack")]
kf = collections.Counter(sum(_fac(r)) for r in AR)
for k in range(6): H_("DepK" + ["Zero", "One", "Two", "Three", "Four", "Five"][k], kf.get(k, 0))
H_("DepFourPlus", sum(v for k, v in kf.items() if k >= 4)); H_("DepGenSrcChan", sum(1 for r in AR if _fac(r)[0] and _fac(r)[1] and _fac(r)[3]))
for i, nm in enumerate(["Gen", "Src", "Dom", "Chan", "Adv"]): H_("DepF" + nm, sum(1 for r in AR if _fac(r)[i]))
gl.append("\\newcommand{\\DeployPlot}{\\begin{tikzpicture}\\let\\xspace\\empty\\begin{axis}[ybar, width=\\columnwidth, height=4.4cm, bar width=18pt, symbolic x coords={0,1,2,3,4,5}, xtick=data, ymin=0, ymax=32, xlabel={Deployment factors tested (of five)}, ylabel={Studies}, label style={font=\\footnotesize}, tick label style={font=\\footnotesize}, nodes near coords, nodes near coords style={font=\\scriptsize}, axis lines*=left, enlarge x limits=0.12]\\addplot[fill=cB!70, draw=cA] coordinates {" + " ".join(f"({k},{kf.get(k, 0)})" for k in range(6)) + "};\\end{axis}\\end{tikzpicture}}")
# evidence family table rows
FT_ = [("Hybrid", "HYBRID"), ("Spatial artifact", "SPATIAL"), ("Frequency", "FREQ"), ("Semantic", "SEMANTIC"), ("Reconstruction", "RECON"), ("Other", None)]
rows_ = []
for nm, key in FT_:
    rs = [r for r in AR if (pfam(r).startswith(key) if key else pfam(r).split()[0] in ("NOISE", "SOURCE", "GEOMETRY"))]
    n_ = len(rs); c = lambda f: sum(1 for r in rs if f(r))
    hvn = c(strong); dd = [ACCd[r["Record_ID"]] for r in rs if r["Record_ID"] in ACCd]
    med_ = f"{_st.median(dd):.1f}" if len(dd) >= 3 else "--"
    rows_.append(f"{nm} & {n_} & {c(lambda r: Y(r, 'Cross_Family_Test'))} & {c(lambda r: Y(r, 'Cross_Dataset_Test'))} & {c(lambda r: Y(r, 'Robustness_Tested'))} & {c(lambda r: Y(r, 'Social_Media_Postprocessing'))} & {c(lambda r: Y(r, 'Adversarial_Attack'))} & {hvn} ({round(100*hvn/n_)}\\%) & {med_} ({len(dd)}) \\\\")
gl.append("\\newcommand{\\FamilyRows}{" + "\n".join(rows_) + "}")
# minimum reporting standard: compliance of the corpus on the checkable items
_op = lambda r: _re.search(r"tpr|fpr|\beer\b|fnr", str(r["Metrics"] or "").lower())
STD = [("R1 Generators and sample sizes reported", lambda r: QAR[r["Record_ID"]][1] == 1.0),
       ("R2 Held-out target, no target data in calibration", lambda r: any(Y(r, h) for h, *_ in DIMS) and r["Target_Data_Used_For_Calibration"] == "NO"),
       ("R3 Cross-family and cross-dataset tested", lambda r: Y(r, "Cross_Family_Test") and Y(r, "Cross_Dataset_Test")),
       ("R4 Threshold-free and operating-point metric", lambda r: bool(_tfm(r)) and bool(_op(r))),
       ("R5 Platform or adversarial test", lambda r: QAR[r["Record_ID"]][5] == 1.0),
       ("R6 Claim supported by own experiments", lambda r: (not clm(r)) or sup(r)),
       ("R7 Code available", lambda r: Y(r, "Code_Available"))]
for i, (nm, fn) in enumerate(STD): H_("Std" + ["One", "Two", "Three", "Four", "Five", "Six", "Seven"][i], sum(1 for r in AR if fn(r)))
met = [sum(1 for _, fn in STD if fn(r)) for r in AR]
H_("StdMedian", int(_st.median(met))); H_("StdAll", sum(1 for m in met if m == 7)); H_("StdAtMostThree", sum(1 for m in met if m <= 3)); H_("StdFiveUp", sum(1 for m in met if m >= 5))
_sl = ["R1 Protocol reported", "R2 Held-out and uncalibrated", "R3 Combined shifts", "R4 Both metric types", "R5 Channel or attack", "R6 Supported claim", "R7 Code released"]
gl.append("\\newcommand{\\StdItemPlot}{" + hbar([(l, sum(1 for r in AR if fn(r))) for l, (_, fn) in zip(_sl, STD)], "Studies meeting the item", 99, "4.2cm") + "}")
mc = collections.Counter(met)
gl.append("\\newcommand{\\StdCountPlot}{\\begin{tikzpicture}\\let\\xspace\\empty\\begin{axis}[ybar, width=\\columnwidth, height=4.2cm, bar width=14pt, symbolic x coords={0,1,2,3,4,5,6,7}, xtick=data, ymin=0, ymax=34, xlabel={Items met (of seven)}, ylabel={Studies}, label style={font=\\footnotesize}, tick label style={font=\\footnotesize}, nodes near coords, nodes near coords style={font=\\scriptsize}, axis lines*=left, enlarge x limits=0.1]\\addplot[fill=cB!70, draw=cA] coordinates {" + " ".join(f"({k},{mc.get(k, 0)})" for k in range(8)) + "};\\end{axis}\\end{tikzpicture}}")
with open(f"{OUT}/results_g_data.tex", "w") as f:
    f.write("% AUTO-GENERATED by scripts/gen_from_workbook.py -- do not edit by hand\n" + "\n".join(gl) + "\n")


# ---------------- verbatim Scopus search strings (supplement) ----------------
_sl = wb["01_Search_Log"]
_hd = [c.value for c in _sl[1]]
for _r in _sl.iter_rows(min_row=2, values_only=True):
    if _r[0] in ("SCOPUS_S01", "SCOPUS_S02"):
        _d = dict(zip(_hd, _r))
        with open(f"{OUT}/search_{_r[0][-3:]}.txt", "w") as _f: _f.write(str(_d["Search_String"]).strip() + "\n")
