# TP de Sécurité

```
Mathieu DARRIBAU
```

## La procédure effectuée :

Tout d'abord on repère toutes les lignes qui pourraient être changées comme par exemple :

- Pour changer le nom d'une variable :

```
Entrée : sWQL = FLQZEsG("gq14V+IwdBHohEVdaK7yS...")

Sortie : query = FLQZEsG("gq14V+IwdBHohEVdaK7yS...")
```

- Pour transformer un string :

```
On récupère la ligne que l'on souhaite modifier puis on afficher son rendu grâce au 'Debug.Print' ce qui donne :

Entrée : Debug.Print FLQZEsG("gq14V+IwdBHohEVdaK7ySOp3F0yC71Zdqra7UKDjUFxotvVUoqbfVaDy8b46X9V")

Sortie : Select * From Win32_NetworkAdapterConfiguration
```

Puis on continue de remplacer les lignes pour la suite du code
Et enfin, une fois cette opération finie, on peut supprimer les quatres fonctions après les diffèrents Sub permettant ainsi de désobfusquer le texte.

## Mais que fait le code ?

Le code permet d'afficher les propriétés des éléments 'root' de Windows

On peut retrouver une URL permettant de télécharger le document contenant une documentation.
