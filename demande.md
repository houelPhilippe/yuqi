---
title: "UC01 - Réception de l'AER, Avis Anticipé d'Export"
abstract-title: |
    L'AER (IE501) est utilisé pour communiquer la mise à l'exportation des marchandises 
au bureau de douane de sortie.
lightbox: true
toc-depth: 5
---

::: {.content-visible when-format="html"}
![SDF-CID-SDS-C1_img05](C:\DEV\_Perso\editeur_python\img\SDF-CID-SDS-C1_img05.jpg){.align-right width=30%}
:::

Nous allons créer une application "Electron/Tauri + Python backend" pour être capable dans un premier temps les fonctionnalités suivantes,

-   Ouvrir / Fermer / Enregistrer des fichiers `.md` et `.qmd`
-   dedededede
-   Barre d’outils avec raccourcis clavier :
    -   Cmd/Ctrl+O : Ouvrir
    -   Cmd/Ctrl+S : Enregistrer
    -   Cmd/Ctrl+W : Fermer
-   Édition et *prévisualisation* Markdown lisible

le [fichier](#titre-3) **doit** être **ouvert** dans un **onglet** séparer pour **être** capable d'ouvrir plusieurs fichiers.

# Titre 1

**Curabitur** ac turpis bibendum, **lacinia** neque a, posuere nisi. Nam quis est nibh. Donec placerat, *mi sit amet laoreet euismod, urna massa lacinia nisl*, nec `hendrerit ipsum nibh vulputate` quam. 

+-------------------------------------+-------------------------------------+
| Col 1                               | Col 2                               |
+:====================================+:===================================:+
| Integer sed enim nec arcu fermentum | Integer sed enim nec arcu fermentum |
+-------------------------------------+-------------------------------------+
| cell                                | cell                                |
+-------------------------------------+-------------------------------------+
: Légende {tbl-colwidths="[32, 37]"}

## Titre 2

: Liste des Traitements. {tbl-colwidths="\[20, 20, 20, 20, 20\]" .striped .bordered}

**Suspendisse** gravida dolor `vitae sapien sollicitudin` gravida. Integer sed enim nec arcu fermentum dapibus sit amet id quam. Nulla ultrices aliquam leo eu semper. Curabitur vulputate eget lectus sed porta. Proin gravida metus eu ipsum posuere laoreet. `Suspendisse` ultrices neque ligula,

```plaintext
MESSAGE
|--- messageSender
|--- Message recipient
|--- Preparation date and time
|--- Message identification
|--- [Message type]
|--- Correlation identifier

```

### Titre 3

*in malesuada massa pharetra egestas. Quisque et cursus magna. Praeent et venenatis neque. Ut quis tempor est. Nulla auctor porta sem sit amet molestie. Quisque vitae ligula consequat, sagittis dolor sed, euismod nunc. Donec neque tellus, posuere quis dui id, fermentum aliquet velit.s*

The transition will be ~~synchronized~~ for all \*\*countries \*\*using a date that will be agreed by ECCG (national applications are expected to manage this data item as a dynamic data element). This technical rule may be replaced by a BRT.

-   in malesuada massa pharetra egestas. Quisque et cursus magna. Praesent et venenatis neque. Ut quis tempor est. Nulla auctor porta sem sit amet molestie. 
-   Quisque vitae ligula consequat, sagittis dolor sed, euismod nunc. 
-   Donec neque tellus, posuere quis dui id, fermentum aliquet velit.

> *The transition will be synchronized for all countries using a date that will be agreed by ECCG (national applications are expected to manage this data item as a dynamic data element).*

+----------+--------------+-----------+
| Colonn1  | Col2         | Col3      |
+==========+:============:+==========:+
| dedede   | dededede     | dedededed |
|          |              |           |
| dededede | dededededede |           |
|          |              |           |
| dedede   |              |           |
+----------+--------------+-----------+
| dedede   | dedede       | dededed   |
+----------+--------------+-----------+
| dededed  | dedede       | dededede  |
+----------+--------------+-----------+
: Liste des Traitements. {tbl-colwidths="[19, 21, 17]"}

The transition will be ~~synchronized~~ for all \*\*countries \*\*using a date that will be agreed by ECCG (national applications are expected to manage this data item as a dynamic data element). This technical rule may be replaced by a BRT.

### Titre de niveau III

1.  The transition will be ~~synchronized~~ for all \*\*countries \*\*using a date that will be agreed by ECCG (national applications are expected to manage this data item as a dynamic data element). This technical rule may be replaced by a BRT.
    
2.  The transition will be ~~synchronized~~ for all \*\*countries \*\*using a date that will be agreed by ECCG (national applications are expected to manage this data item as a dynamic data element). This technical rule may be replaced by a BRT.
    
3.  The transition will be ~~synchronized~~ for all \*\*countries \*\*using a date that will be agreed by ECCG (national applications are expected to manage this data item as a dynamic data element). This technical rule may be replaced by a BRT.