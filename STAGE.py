import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from sklearn.linear_model import LinearRegression

from sklearn.model_selection import train_test_split
from sklearn.linear_model import LinearRegression
from sklearn.metrics import mean_squared_error , r2_score

#data = "STAGE.xlsx"

stage = pd.read_excel("STAGE.xlsx")

print(stage)
print(stage.head())
print(stage.info())

print(stage.columns)

print(stage.describe(include='all'))

print(stage.isnull().sum())

#VENTE PAR PAYS

#calcul de vente
vente_pays = stage.groupby('Pays_vente')['Prix_vente (€)'].sum().reset_index()
print(vente_pays)

#ici trie dans l'ordre decroissant

vente_trie = vente_pays.sort_values(by = ['Prix_vente (€)']  , ascending=False)

print(vente_trie)

top_marques = stage['Marque'].value_counts()


#Calcul du chiffre d'affaire
chiffre_affaire_total = stage['Prix_vente (€)'].sum()
print(f"Chiffre d'affaires total: {chiffre_affaire_total:,.2f}€")


#Prix moyen des ventes :

prix_moyen = stage['Prix_vente (€)'].mean()
print(f"Prix moyen de vente: {prix_moyen:,.2f}€")

#Moyenne des remise :

remise_moyenne = stage['Remise (€)'].mean()
print(f"Remise moyenne: {remise_moyenne:,.2f}€")

#score moyen satisfaction

satisfaction_moyenne = stage['Score_satisfaction_client (1-10)'].mean()
print(f"Score moyen de satisfaction: {satisfaction_moyenne:.2f}/10")

#delaie moyen de livraison :

delai_moyen = stage['Délai_livraison (jours)'].mean()
print(f"Délai moyen de livraison: {delai_moyen:.1f} jours")

#ANALYSE DES VENTES PAR VEHICULE

ventes_par_modele = stage.groupby(['Marque', 'Modèle']).agg({
    'ID_vente': 'count',
    'Prix_vente (€)': ['mean', 'sum']
}).sort_values(('ID_vente', 'count'), ascending=False)

ventes_par_modele.columns = ['Nombre_ventes', 'Prix_moyen', 'chiffre_affaire_total']
print(ventes_par_modele.head(10))

# Visualisation des ventes par modèle
top10_modeles = ventes_par_modele.head(10)
top10_modeles['Nombre_ventes'].plot(kind='bar', figsize=(12,6))
plt.title('Top 10 des modèles les plus vendus')
plt.ylabel('Nombre de ventes')
plt.xticks(rotation=45)
plt.show()


#VENTE PAR MARQUE

vente_marques = stage['Marque'].value_counts()

# Trier par nombre de ventes (descendant)
ventes_par_marque = vente_marques.sort_values()

plt.figure(figsize=(12, 8))
bars = plt.barh(ventes_par_marque.index, ventes_par_marque.values, color='skyblue')

for bar in bars:
    width = bar.get_width()
    plt.text(width + 0.5,
             bar.get_y() + bar.get_height()/2,
             f'{int(width)} ({width/len(stage)*100:.1f}%)',  # Ajoute le pourcentage
             va='center', ha='left')

plt.title('Nombre de ventes par marque', pad=20)
plt.xlabel('Nombre de ventes')
plt.ylabel('Marque')
plt.grid(axis='x', linestyle='--', alpha=0.7)
plt.tight_layout()
plt.show()


#Vente par années

stage['Année_vente'] = stage['Date_vente'].dt.year
evolution_annuelle = stage.groupby('Année_vente').agg({
    'ID_vente': 'count',
    'Prix_vente (€)': 'sum'
}).rename(columns={'ID_vente': 'Nombre_ventes', 'Prix_vente (€)': 'CA'})

# Formatage des nombres pour les étiquettes
def format_number(x):
    if x >= 1e6:
        return f"{x/1e6:.1f}M"
    elif x >= 1e3:
        return f"{x/1e3:.1f}K"
    return str(x)

# Visualisation avec étiquettes
fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(12, 12))

#  Nombre de ventes
bars1 = ax1.bar(evolution_annuelle.index.astype(str), 
                evolution_annuelle['Nombre_ventes'], 
                color='skyblue')
ax1.set_title('Nombre de ventes par année', pad=20, fontsize=14)
ax1.set_ylabel('Nombre de ventes', fontsize=12)
ax1.grid(axis='y', linestyle='--', alpha=0.7)

# Ajoutons  des étiquettes sur les barres
for bar in bars1:
    height = bar.get_height()
    ax1.text(bar.get_x() + bar.get_width()/2., height,
             format_number(int(height)),
             ha='center', va='bottom', fontsize=10)

#  Chiffre d'affaires
bars2 = ax2.bar(evolution_annuelle.index.astype(str), 
                evolution_annuelle['CA'], 
                color='salmon')
ax2.set_title('Chiffre d\'affaires par année', pad=20, fontsize=14)
ax2.set_ylabel('CA (€)', fontsize=12)
ax2.grid(axis='y', linestyle='--', alpha=0.7)

# Ajout des étiquettes sur les barres
for bar in bars2:
    height = bar.get_height()
    ax2.text(bar.get_x() + bar.get_width()/2., height,
             f"{height/1e6:.1f}M€" if height >= 1e6 else f"{height/1e3:.0f}K€",
             ha='center', va='bottom', fontsize=10)

plt.tight_layout()
plt.show()

# Affichage des données sous forme de tableau
print("\nTableau récapitulatif par année:")
print(evolution_annuelle)


#ANALYSE CROISE : client marque carburant


# Analyse croisée complète
analyse_croisee = stage.groupby(['Marque', 'Année_modèle', 'Type_carburant']).agg({
    'ID_vente': 'count',
    'Prix_vente (€)': 'mean',
    'Client_type': lambda x: x.value_counts().index[0],
    'Score_satisfaction_client (1-10)': 'mean'
}).rename(columns={
    'ID_vente': 'Nombre_ventes',
    'Prix_vente (€)': 'Prix_moyen',
    'Score_satisfaction_client (1-10)': 'Satisfaction_moyenne'
}).reset_index()

# 1. Top 10 des combinaisons les plus vendues
top10_combinaisons = analyse_croisee.sort_values('Nombre_ventes', ascending=False).head(10)

plt.figure(figsize=(14, 8))
sns.barplot(data=top10_combinaisons, 
            x='Nombre_ventes', 
            y='Marque', 
            hue='Type_carburant',
            palette='viridis')

# Ajout des étiquettes de données
for index, row in top10_combinaisons.iterrows():
    plt.text(row['Nombre_ventes'], 
             index % len(top10_combinaisons['Marque'].unique()), 
             f"{row['Nombre_ventes']}\n{row['Prix_moyen']:.0f}€",
             va='center', ha='left')

plt.title('Top 10 des combinaisons Marque/Carburant les plus vendues')
plt.xlabel('Nombre de ventes')
plt.ylabel('Marque')
plt.legend(title='Type carburant', bbox_to_anchor=(1.05, 1), loc='upper left')
plt.tight_layout()
plt.show()

# 2. Heatmap pour toutes les marques
pivot_table = pd.pivot_table(analyse_croisee,
                            values='Nombre_ventes',
                            index='Marque',
                            columns='Type_carburant',
                            aggfunc='sum',
                            fill_value=0)

plt.figure(figsize=(16, 10))
sns.heatmap(pivot_table, 
            annot=True, 
            fmt='g', 
            cmap='YlGnBu',
            linewidths=.5)
plt.title('Répartition des ventes par marque et type de carburant')
plt.xlabel('Type de carburant')
plt.ylabel('Marque')
plt.show()

# 3. Visualisation par marque avec subplots
# Version corrigée avec gestion du warning
plt.figure(figsize=(12, 6))
top_marques = stage['Marque'].value_counts().head(10)

# Correction : utilisation de hue au lieu de palette
ax = sns.barplot(x=top_marques.values, 
                 y=top_marques.index, 
                 hue=top_marques.index,  # Ajout de hue
                 palette="Blues_d", 
                 dodge=False, 
                 legend=False)

# Personnalisation esthétique
ax.bar_label(ax.containers[0], fmt='%g', padding=3)  # Étiquettes de données

plt.title('Top 10 des marques les plus vendues', pad=20, fontsize=14)
plt.xlabel('Nombre de ventes', fontsize=12)
plt.ylabel('')
plt.grid(axis='x', alpha=0.2)

# Suppression des bordures superflues
sns.despine(left=True, bottom=True)
plt.tight_layout()
plt.show()

# 2. Répartition carburant pour les 5 premières marques (version simplifiée)
top5_marques = top_marques.head(5).index

plt.figure(figsize=(10, 6))
for i, marque in enumerate(top5_marques):
    subset = stage[stage['Marque'] == marque]
    counts = subset['Type_carburant'].value_counts(normalize=True)
    plt.barh(marque, counts.values[0], color='skyblue', label=counts.index[0] if i == 0 else "")
    plt.barh(marque, counts.values[1], left=counts.values[0], color='salmon', label=counts.index[1] if i == 0 else "")

plt.title('Répartition carburant pour les 5 premières marques')
plt.xlabel('Proportion')
plt.legend(title='Type carburant', bbox_to_anchor=(1.05, 1))
sns.despine()
plt.tight_layout()
plt.show()


#4. Analyse des clients par marque
client_marque = stage.groupby(['Marque', 'Client_type']).size().unstack()

client_marque.plot(kind='bar', 
                  stacked=True, 
                  figsize=(14, 8),
                  colormap='Paired')

plt.title('Répartition des types de clients par marque')
plt.xlabel('Marque')
plt.ylabel('Nombre de clients')
plt.xticks(rotation=45)
plt.legend(title='Type de client', bbox_to_anchor=(1.05, 1), loc='upper left')

# Ajout des pourcentages
for n, x in enumerate([*client_marque.index.values]):
    for (proportion, y_loc) in zip(client_marque.loc[x],
                                  client_marque.loc[x].cumsum()):
        plt.text(x=n - 0.17,
                y=(y_loc - proportion) + proportion / 2,
                s=f'{proportion}\n({np.round(proportion/ client_marque.loc[x].sum()*100, 1)}%)',
                color="black",
                fontsize=8,
                fontweight="bold")

plt.tight_layout()
plt.show()

#matrice de correlation :

numerical_cols = ['Prix_vente (€)', 'Année_modèle', 'Kilométrage (km)', 'Garantie (mois)', 
                 'Score_satisfaction_client (1-10)', 'Remise (€)', 'Délai_livraison (jours)']

plt.figure(figsize=(12, 8))
sns.heatmap(stage[numerical_cols].corr(), annot=True, cmap='coolwarm', center=0)
plt.title('Matrice de corrélation')
plt.show()



top_marques = stage['Marque'].value_counts().head(10)
#analyse des prix par marque et type de carburant
plt.figure(figsize=(14, 8))
sns.boxplot(x='Marque', y='Prix_vente (€)', hue='Type_carburant', data=stage[stage['Marque'].isin(top_marques.index)])
plt.xticks(rotation=45)
plt.title('Prix moyen par marque et type de carburant')
plt.show()

