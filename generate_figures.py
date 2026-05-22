"""Generate all figures for the report."""
import warnings
warnings.filterwarnings('ignore')

from pathlib import Path
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import seaborn as sns

from sklearn.model_selection import train_test_split
from sklearn.impute import SimpleImputer
from sklearn.pipeline import Pipeline
from sklearn.ensemble import RandomForestClassifier
from sklearn.metrics import (
    accuracy_score, balanced_accuracy_score, precision_score, recall_score,
    f1_score, roc_auc_score, confusion_matrix, precision_recall_curve,
)
from sklearn.inspection import permutation_importance

RANDOM_SEED = 42
np.random.seed(RANDOM_SEED)
plt.style.use('seaborn-v0_8-whitegrid')
sns.set_context('notebook')
BLUE   = '#2E75B6'
ORANGE = '#C55A11'
GREEN  = '#375623'
GRAY   = '#7F7F7F'

FIGURES_DIR = Path(r'C:\DAVID DANHOUEGNON ECC\OneDrive\DOC DE COURS 2A\Machine Learning\MiniProject_When_ML_Fails_SECOM\figures')
FIGURES_DIR.mkdir(exist_ok=True)

DATA_PATH = Path(r'C:\DAVID DANHOUEGNON ECC\OneDrive\DOC DE COURS 2A\Machine Learning\uci-secom data set ML.csv')

# ── Data loading ─────────────────────────────────────────────────────────────
df_raw = pd.read_csv(DATA_PATH)
print(f'Dataset: {df_raw.shape}')

def prepare_secom_features(df, missing_threshold=0.50):
    df = df.copy()
    y = (df['Pass/Fail'] == 1).astype(int)
    X = df.drop(columns=['Pass/Fail'])
    if 'Time' in X.columns:
        X = X.drop(columns=['Time'])
    high_missing = X.isna().mean()[lambda s: s > missing_threshold].index.tolist()
    X = X.drop(columns=high_missing)
    X = pd.get_dummies(X, columns=['Phase'], dummy_na=False)
    constant_cols = X.nunique(dropna=True)[lambda s: s <= 1].index.tolist()
    X = X.drop(columns=constant_cols)
    return X, y

X_clean, y = prepare_secom_features(df_raw)
X_train_clean, X_test_clean, y_train, y_test = train_test_split(
    X_clean, y, test_size=0.30, random_state=RANDOM_SEED, stratify=y)

def build_rf(seed=RANDOM_SEED, class_weight='balanced_subsample'):
    return Pipeline([
        ('imputer', SimpleImputer(strategy='median')),
        ('model', RandomForestClassifier(
            n_estimators=400, min_samples_leaf=2,
            class_weight=class_weight, random_state=seed, n_jobs=-1)),
    ])

def eval_model(model, X, y_true, label):
    y_pred  = model.predict(X)
    y_proba = model.predict_proba(X)[:, 1]
    return dict(setting=label,
                accuracy=accuracy_score(y_true, y_pred),
                balanced_accuracy=balanced_accuracy_score(y_true, y_pred),
                precision_fail=precision_score(y_true, y_pred, zero_division=0),
                recall_fail=recall_score(y_true, y_pred, zero_division=0),
                f1_fail=f1_score(y_true, y_pred, zero_division=0),
                roc_auc=roc_auc_score(y_true, y_proba))

def make_shortcut(y_vals, reliability, seed, mode='aligned'):
    rng = np.random.default_rng(seed)
    y_vals = np.asarray(y_vals).astype(int)
    if mode == 'random':
        return rng.integers(0, 2, len(y_vals))
    sc = y_vals.copy() if mode == 'aligned' else 1 - y_vals.copy()
    flip = rng.random(len(y_vals)) > reliability
    sc[flip] = 1 - sc[flip]
    return sc

def add_shortcut(X, y_vals, reliability, seed, mode):
    Xn = X.copy()
    Xn['shortcut_alarm'] = make_shortcut(y_vals, reliability, seed, mode)
    return Xn

RELIABILITY = 0.98

# ── Reference model ───────────────────────────────────────────────────────────
ref_model = build_rf()
ref_model.fit(X_train_clean, y_train)

# ── Broken / corrected pipelines ─────────────────────────────────────────────
Xtr_sc  = add_shortcut(X_train_clean, y_train,  RELIABILITY, RANDOM_SEED+1, 'aligned')
Xte_id  = add_shortcut(X_test_clean,  y_test,   RELIABILITY, RANDOM_SEED+2, 'aligned')
Xte_rnd = add_shortcut(X_test_clean,  y_test,   RELIABILITY, RANDOM_SEED+3, 'random')
Xte_ood = add_shortcut(X_test_clean,  y_test,   RELIABILITY, RANDOM_SEED+4, 'inverted')

broken_model = build_rf()
broken_model.fit(Xtr_sc, y_train)

Xtr_san = Xtr_sc.drop(columns=['shortcut_alarm'])
Xte_san = Xte_ood.drop(columns=['shortcut_alarm'])
corrected_model = build_rf()
corrected_model.fit(Xtr_san, y_train)

# ── Figure 1 — ROC-AUC comparison ────────────────────────────────────────────
summary = pd.DataFrame([
    eval_model(broken_model,    Xte_id,          y_test, 'Broken\n(shortcut aligné)'),
    eval_model(broken_model,    Xte_ood,         y_test, 'Broken\n(shortcut inversé)'),
    eval_model(corrected_model, Xte_san,         y_test, 'Corrigé\n(shortcut retiré)'),
    eval_model(ref_model,       X_test_clean,    y_test, 'Référence\n(sans shortcut)'),
])

fig, ax = plt.subplots(figsize=(8, 5))
colors = [ORANGE, '#E74C3C', GREEN, BLUE]
bars = ax.bar(summary['setting'], summary['roc_auc'], color=colors, width=0.55, edgecolor='white')
for bar, val in zip(bars, summary['roc_auc']):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.01,
            f'{val:.4f}', ha='center', va='bottom', fontsize=11, fontweight='bold')
ax.set_ylim(0, 1.08)
ax.set_ylabel('ROC-AUC', fontsize=12)
ax.set_title('ROC-AUC selon le contexte d\'évaluation', fontsize=13, fontweight='bold')
ax.axhline(0.5, linestyle='--', color='gray', linewidth=0.8, alpha=0.6)
ax.text(3.55, 0.515, 'aléatoire', fontsize=9, color='gray', va='bottom')
plt.tight_layout()
plt.savefig(FIGURES_DIR / 'fig1_roc_comparison.png', dpi=150, bbox_inches='tight')
plt.close()
print('fig1 done')

# ── Figure 2 — Sensitivity analysis ──────────────────────────────────────────
reliability_levels = [0.55, 0.70, 0.85, 0.98]
rows = []
for i, rel in enumerate(reliability_levels):
    Xtr = add_shortcut(X_train_clean, y_train, rel, RANDOM_SEED+100+i, 'aligned')
    Xid = add_shortcut(X_test_clean,  y_test,  rel, RANDOM_SEED+200+i, 'aligned')
    Xod = add_shortcut(X_test_clean,  y_test,  rel, RANDOM_SEED+300+i, 'inverted')
    m = build_rf(seed=RANDOM_SEED+i)
    m.fit(Xtr, y_train)
    for ctx, Xe in [('In-distribution (aligné)', Xid), ('Out-of-distribution (inversé)', Xod)]:
        r = eval_model(m, Xe, y_test, f'{rel}')
        r['reliability'] = rel
        r['context'] = ctx
        rows.append(r)
sens_df = pd.DataFrame(rows)

fig, axes = plt.subplots(1, 2, figsize=(12, 4.5))
palette = {'In-distribution (aligné)': BLUE, 'Out-of-distribution (inversé)': ORANGE}
for ax, metric, ylabel in zip(axes, ['roc_auc', 'f1_fail'], ['ROC-AUC', 'F1 (classe Fail)']):
    for ctx, grp in sens_df.groupby('context'):
        ax.plot(grp['reliability'], grp[metric], marker='o', label=ctx,
                color=palette[ctx], linewidth=2, markersize=7)
    ax.set_xlabel('Fiabilité du shortcut', fontsize=11)
    ax.set_ylabel(ylabel, fontsize=11)
    ax.set_ylim(-0.05, 1.08)
    ax.set_xticks(reliability_levels)
    ax.legend(fontsize=9)
    ax.set_title(f'{ylabel} selon la fiabilité', fontsize=12, fontweight='bold')
plt.suptitle('Impact de la fiabilité du shortcut sur les performances ID vs OOD', fontsize=12)
plt.tight_layout()
plt.savefig(FIGURES_DIR / 'fig2_sensitivity.png', dpi=150, bbox_inches='tight')
plt.close()
print('fig2 done')

# ── Figure 3 — Feature importance ────────────────────────────────────────────
rf_broken = broken_model.named_steps['model']
native_imp = pd.DataFrame({
    'feature': Xtr_sc.columns,
    'importance': rf_broken.feature_importances_,
}).sort_values('importance', ascending=False).head(12)

perm = permutation_importance(broken_model, Xte_id, y_test,
                               scoring='roc_auc', n_repeats=10,
                               random_state=RANDOM_SEED, n_jobs=-1)
perm_imp = pd.DataFrame({
    'feature': Xte_id.columns,
    'importance': perm.importances_mean,
    'std': perm.importances_std,
}).sort_values('importance', ascending=False).head(12)

fig, axes = plt.subplots(1, 2, figsize=(13, 5))
colors_native = [ORANGE if f == 'shortcut_alarm' else BLUE for f in native_imp['feature']]
axes[0].barh(native_imp['feature'][::-1], native_imp['importance'][::-1],
             color=colors_native[::-1], edgecolor='white')
axes[0].set_title('Importance native (MDI)', fontsize=12, fontweight='bold')
axes[0].set_xlabel('Importance')
axes[0].axvline(0, color='gray', linewidth=0.5)

colors_perm = [ORANGE if f == 'shortcut_alarm' else GREEN for f in perm_imp['feature']]
axes[1].barh(perm_imp['feature'][::-1], perm_imp['importance'][::-1],
             xerr=perm_imp['std'][::-1], color=colors_perm[::-1],
             edgecolor='white', capsize=3)
axes[1].set_title('Permutation importance (ROC-AUC)', fontsize=12, fontweight='bold')
axes[1].set_xlabel('Diminution ROC-AUC (permutation)')

from matplotlib.patches import Patch
legend_els = [Patch(color=ORANGE, label='shortcut_alarm'), Patch(color=BLUE, label='autres features')]
fig.legend(handles=legend_els, loc='lower center', ncol=2, fontsize=10, bbox_to_anchor=(0.5, -0.02))
plt.suptitle('shortcut_alarm domine massivement l\'importance — modèle cassé', fontsize=12)
plt.tight_layout(rect=[0, 0.04, 1, 1])
plt.savefig(FIGURES_DIR / 'fig3_feature_importance.png', dpi=150, bbox_inches='tight')
plt.close()
print('fig3 done')

# ── Figure 4 — Confusion matrices shortcut ───────────────────────────────────
def plot_cm(y_true, y_pred, title, ax):
    cm = confusion_matrix(y_true, y_pred, labels=[0, 1])
    sns.heatmap(cm, annot=True, fmt='d', cmap='Blues', ax=ax,
                xticklabels=['Pass', 'Fail'], yticklabels=['Pass', 'Fail'],
                cbar=False, annot_kws={'size': 13})
    ax.set_title(title, fontsize=11, fontweight='bold')
    ax.set_xlabel('Prédiction'); ax.set_ylabel('Vérité')

fig, axes = plt.subplots(1, 2, figsize=(9, 4))
plot_cm(y_test, broken_model.predict(Xte_id),  'Broken — shortcut aligné',  axes[0])
plot_cm(y_test, broken_model.predict(Xte_ood), 'Broken — shortcut inversé', axes[1])
plt.suptitle('Le même modèle, les mêmes lignes — seule la corrélation du shortcut change',
             fontsize=11)
plt.tight_layout()
plt.savefig(FIGURES_DIR / 'fig4_confusion_shortcut.png', dpi=150, bbox_inches='tight')
plt.close()
print('fig4 done')

# ── Figure 5 — Multi-seed validation ─────────────────────────────────────────
def run_one_seed(seed):
    Xtr, Xte, ytr, yte = train_test_split(X_clean, y, test_size=0.30,
                                           random_state=seed, stratify=y)
    Xtr_s = add_shortcut(Xtr, ytr, RELIABILITY, seed+1, 'aligned')
    Xte_i = add_shortcut(Xte, yte, RELIABILITY, seed+2, 'aligned')
    Xte_o = add_shortcut(Xte, yte, RELIABILITY, seed+3, 'inverted')
    brk = build_rf(seed); brk.fit(Xtr_s, ytr)
    crt = build_rf(seed); crt.fit(Xtr,   ytr)
    rows_s = []
    for lbl, Xe, ye in [('Broken ID', Xte_i, yte), ('Broken OOD', Xte_o, yte),
                         ('Corrigé',   Xte,   yte)]:
        m = brk if 'Broken' in lbl else crt
        r = eval_model(m, Xe, ye, lbl)
        r['seed'] = seed
        rows_s.append(r)
    return pd.DataFrame(rows_s)

seed_df = pd.concat([run_one_seed(s) for s in [7, 11, 19, 23, 42]], ignore_index=True)
seed_summary = seed_df.groupby('setting')['roc_auc'].agg(['mean', 'std']).reset_index()
seed_summary.columns = ['setting', 'mean', 'std']

fig, ax = plt.subplots(figsize=(7, 4.5))
order = ['Broken ID', 'Corrigé', 'Broken OOD']
palette_s = {'Broken ID': ORANGE, 'Broken OOD': '#E74C3C', 'Corrigé': GREEN}
for _, row in seed_summary.iterrows():
    if row['setting'] not in order: continue
    ax.bar(row['setting'], row['mean'], color=palette_s[row['setting']],
           width=0.5, edgecolor='white')
    ax.errorbar(row['setting'], row['mean'], yerr=row['std'],
                fmt='none', color='black', capsize=6, linewidth=2)
    ax.text(order.index(row['setting']), row['mean'] + row['std'] + 0.015,
            f"{row['mean']:.4f}\n±{row['std']:.4f}", ha='center', fontsize=10)
ax.set_ylim(0, 1.1)
ax.set_ylabel('ROC-AUC (moyenne ± écart-type, 5 seeds)', fontsize=11)
ax.set_title('Robustesse de l\'effet sur 5 splits aléatoires', fontsize=12, fontweight='bold')
plt.tight_layout()
plt.savefig(FIGURES_DIR / 'fig5_multiseed.png', dpi=150, bbox_inches='tight')
plt.close()
print('fig5 done')

# ── Figure 6 — Class imbalance ────────────────────────────────────────────────
Xtr_in, Xva_in, ytr_in, yva_in = train_test_split(
    X_train_clean, y_train, test_size=0.25, random_state=RANDOM_SEED, stratify=y_train)

imb_brk = build_rf(class_weight=None);    imb_brk.fit(Xtr_in, ytr_in)
imb_crt = build_rf(class_weight='balanced_subsample'); imb_crt.fit(Xtr_in, ytr_in)
y_val_proba = imb_crt.predict_proba(Xva_in)[:, 1]
prec, rec, thr = precision_recall_curve(yva_in, y_val_proba)
f1s = 2*prec*rec/(prec+rec+1e-12)
best_thr = thr[int(np.nanargmax(f1s[:-1]))]

y_pred_brk_imb = imb_brk.predict(X_test_clean)
y_proba_crt_imb = imb_crt.predict_proba(X_test_clean)[:, 1]
y_pred_crt_imb = (y_proba_crt_imb >= best_thr).astype(int)

fig, axes = plt.subplots(1, 2, figsize=(9, 4))
plot_cm(y_test, y_pred_brk_imb,
        f'Cassé\n(sans class_weight, seuil 0.5)', axes[0])
plot_cm(y_test, y_pred_crt_imb,
        f'Corrigé\n(class_weight + seuil {best_thr:.3f})', axes[1])
plt.suptitle('Class imbalance — le modèle non pondéré rate 100% des défauts', fontsize=11)
plt.tight_layout()
plt.savefig(FIGURES_DIR / 'fig6_confusion_imbalance.png', dpi=150, bbox_inches='tight')
plt.close()
print('fig6 done')

# ── Figure 7 — Distribution shift ────────────────────────────────────────────
phase_fail = pd.crosstab(df_raw['Phase'], df_raw['Pass/Fail'])
eligible = phase_fail[(phase_fail[-1] >= 10) & (phase_fail[1] >= 4)]
holdout_phase = eligible[1].idxmax()
print(f'Holdout phase: {holdout_phase}')

phase_series = df_raw.loc[X_clean.index, 'Phase']
src_mask = phase_series != holdout_phase
tgt_mask = phase_series == holdout_phase
Xs, ys = X_clean.loc[src_mask], y.loc[src_mask]
Xt, yt = X_clean.loc[tgt_mask], y.loc[tgt_mask]

Xstr, Xste, ystr, yste = train_test_split(Xs, ys, test_size=0.30, random_state=RANDOM_SEED, stratify=ys)
Xtca, Xtte, ytca, ytte = train_test_split(Xt, yt, test_size=0.70, random_state=RANDOM_SEED, stratify=yt)

shf_brk = build_rf(); shf_brk.fit(Xstr, ystr)
shf_crt = build_rf(); shf_crt.fit(pd.concat([Xstr, Xtca]), pd.concat([ystr, ytca]))

shift_data = {
    'Broken\n(source)':   eval_model(shf_brk, Xste, yste, '')['roc_auc'],
    'Broken\n(target)':   eval_model(shf_brk, Xtte, ytte, '')['roc_auc'],
    'Corrigé\n(source)':  eval_model(shf_crt, Xste, yste, '')['roc_auc'],
    'Corrigé\n(target)':  eval_model(shf_crt, Xtte, ytte, '')['roc_auc'],
}

fig, ax = plt.subplots(figsize=(8, 4.5))
bar_colors = [ORANGE, '#E74C3C', GREEN, '#27AE60']
bars = ax.bar(shift_data.keys(), shift_data.values(), color=bar_colors, width=0.55, edgecolor='white')
for bar, val in zip(bars, shift_data.values()):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.01,
            f'{val:.4f}', ha='center', va='bottom', fontsize=11, fontweight='bold')
ax.set_ylim(0, 1.05)
ax.set_ylabel('ROC-AUC', fontsize=11)
ax.set_title(f'Distribution shift — phase «{holdout_phase}» tenue hors entraînement\n'
             f'Correction : inclusion de données de la phase cible en calibration',
             fontsize=11, fontweight='bold')
plt.tight_layout()
plt.savefig(FIGURES_DIR / 'fig7_distribution_shift.png', dpi=150, bbox_inches='tight')
plt.close()
print('fig7 done')

print('\nAll figures saved to:', FIGURES_DIR)
print('Files:', [f.name for f in sorted(FIGURES_DIR.glob('*.png'))])
