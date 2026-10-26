'use strict'; // Active le mode strict de JavaScript pour détecter certaines erreurs.

 /* =============================================================================
    Cable List Builder – IO LIST -> CABLE LIST (EQUIPMENT sheet)
    Ce programme transforme une liste d'entrées/sorties Excel en liste de câbles.
    Chaque boucle commence au FACP et revient au FACP.
    Les appareils conservent l'ordre dans lequel ils apparaissent dans la IO LIST.
    Les appareils sans adresse de boucle sont raccordés en dérivation.
    ============================================================================= */

// ---------- Paramètres par défaut modifiables dans l'interface ----------

const DEFAULTS = { // Regroupe les valeurs utilisées lorsque l'utilisateur ne modifie pas les options.
  cableType: '2x1,5mmq', // Définit le type de câble utilisé par défaut.
  shield: 'SH', // Définit la valeur par défaut de la colonne SHIELD.
  armored: '-', // Définit la valeur par défaut de la colonne ARMORED.
  facpName: 'FACP', // Définit le nom du panneau central d'alarme incendie.
  facpDescription: '', // Définit la description par défaut du FACP.
  removeText: ['-PORTA TAGLIAFUOCO'], // Liste des textes à supprimer des descriptions.
  hubPrefixes: ['CCP'], // Définit les préfixes identifiés comme concentrateurs ou hubs.
  addressMode: 'io', // Utilise l'adresse IO LIST ; 'type' utilise la table TYPE_ADDRESS.
}; // Termine la définition des paramètres par défaut.

const TYPE_ADDRESS = { // Associe certains types d'appareils à des adresses prédéfinies.
  MCP: '004', // Définit l'adresse utilisée pour les appareils MCP en mode type.
  MCPO: '017', // Définit l'adresse utilisée pour les appareils MCPO.
  ML: '017', // Définit l'adresse utilisée pour les appareils ML.
  ALM: '013', // Définit l'adresse utilisée pour les appareils ALM.
  MIO: '014', // Définit l'adresse utilisée pour les appareils MIO.
  SD: '015', // Définit l'adresse utilisée pour les appareils SD.
  ASD: '003', // Définit l'adresse utilisée pour les appareils ASD.
  SOA: '002', // Définit l'adresse utilisée pour les appareils SOA.
}; // Termine la table des adresses par type.

const HEADER_ALIASES = { // Définit les noms possibles des colonnes dans le fichier source.
  loop: ['LOOP'], // Reconnaît la colonne contenant le numéro de boucle.
  prefix: ['PREFISSO', 'PREFIX'], // Reconnaît le préfixe de l'appareil en italien ou en anglais.
  number: ["NUMERO DELL'APPARECCHIATURA", 'NUMERO', 'NUMBER'], // Reconnaît le numéro de l'appareil.
  address: ['INDIRIZZO', 'ADDRESS'], // Reconnaît la colonne de l'adresse.
  desc: ['DESCRIZIONE', 'DESCRIPTION'], // Reconnaît la colonne de description.
  area: ['AREA', 'ZONE'], // Reconnaît la colonne de zone ou de secteur.
}; // Termine la définition des alias de colonnes.

// =============================================================================
// 1. LECTURE DE LA IO LIST : fonctions indépendantes de l'interface graphique.
// =============================================================================

const norm = (v) => // Déclare une fonction qui normalise une valeur pour comparer les en-têtes.
  v == null ? '' : String(v).replace(/\s+/g, ' ').trim().toUpperCase(); // Convertit en texte, uniformise les espaces et passe en majuscules.

function clean(v) { // Déclare une fonction qui nettoie les valeurs lues dans Excel.
  const s = v == null ? '' : String(v).trim(); // Convertit la valeur en texte et supprime les espaces aux extrémités.
  return s === '-' || s === '--' ? '' : s; // Remplace les tirets seuls par une chaîne vide.
} // Termine la fonction de nettoyage.

const escapeRe = (s) => // Déclare une fonction qui protège les caractères spéciaux des expressions régulières.
  s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); // Ajoute une barre oblique inverse devant les caractères spéciaux.

// Convertit une cellule ExcelJS en valeur simple exploitable par le programme.
function cellText(v) { // Déclare la fonction de conversion des cellules Excel.
  if (v == null) return null; // Retourne null si la cellule est vide ou non définie.

  if (v instanceof Date) return v.toISOString(); // Convertit une date en texte au format ISO.

  if (typeof v === 'object') { // Traite les cellules contenant des objets ExcelJS.
    if (v.error) return null; // Ignore les cellules qui contiennent une erreur Excel.

    if ('result' in v) return cellText(v.result); // Récupère le résultat d'une formule Excel, puis le convertit.

    if (v.richText) { // Vérifie si la cellule contient du texte enrichi.
      return v.richText.map((t) => t.text).join(''); // Assemble les fragments de texte enrichi.
    } // Termine le traitement du texte enrichi.

    if (v.text !== undefined) return cellText(v.text); // Utilise la propriété text si elle existe.

    return String(v); // Convertit en texte les autres objets non reconnus.
  } // Termine le traitement des objets.

  return v; // Retourne directement les valeurs simples : texte, nombre ou booléen.
} // Termine la conversion des cellules Excel.

// Recherche la ligne des en-têtes et identifie les colonnes importantes.
function findHeader(matrix) { // Reçoit une matrice contenant toutes les lignes de la feuille Excel.
  const limit = Math.min(matrix.length, 15); // Limite la recherche aux 15 premières lignes.

  for (let r = 0; r < limit; r++) { // Parcourt les lignes susceptibles de contenir les en-têtes.
    const heads = (matrix[r] || []).map(norm); // Normalise les cellules de la ligne courante.
    const cols = {}; // Prépare un objet qui stockera les indices des colonnes trouvées.

    for (const [field, aliases] of Object.entries(HEADER_ALIASES)) { // Parcourt les champs recherchés et leurs alias.
      let idx = -1; // Initialise l'indice de colonne à -1, valeur indiquant une absence.

      // Recherche une correspondance exacte entre l'en-tête et un alias.
      for (const a of aliases) { // Parcourt les noms alternatifs du champ.
        idx = heads.indexOf(a); // Recherche l'indice exact de l'alias dans la ligne.
        if (idx >= 0) break; // Arrête la recherche dès qu'une correspondance est trouvée.
      } // Termine la recherche exacte.

      // Si aucun résultat exact n'existe, recherche une correspondance partielle.
      if (idx < 0) { // Exécute cette recherche uniquement si l'alias exact est absent.
        for (const a of aliases) { // Parcourt les alias disponibles.
          idx = heads.findIndex((h) => h.startsWith(a)); // Cherche un en-tête commençant par l'alias.
          if (idx >= 0) break; // Arrête la recherche dès qu'un en-tête correspond.
        } // Termine la recherche partielle.
      } // Termine la condition de recherche partielle.

      if (idx >= 0) cols[field] = idx; // Enregistre l'indice de la colonne lorsqu'elle a été trouvée.
    } // Termine la recherche de tous les champs.

    if ('loop' in cols && 'prefix' in cols && 'number' in cols) { // Vérifie que les trois colonnes obligatoires existent.
      return { row: r, cols }; // Retourne l'indice de la ligne d'en-tête et les indices des colonnes.
    } // Termine la vérification des colonnes obligatoires.
  } // Termine le parcours des 15 premières lignes.

  return null; // Signale qu'aucune ligne d'en-tête valide n'a été trouvée.
} // Termine la fonction de recherche des en-têtes.

// Transforme les lignes Excel en objets représentant les appareils.
function parseIoMatrix(matrix, opts) { // Reçoit la matrice Excel et les options de traitement.
  const hdr = findHeader(matrix); // Recherche la ligne d'en-tête et les indices de colonnes.

  if (!hdr) return null; // Retourne null si les colonnes obligatoires ne sont pas trouvées.

  const devices = []; // Prépare la liste des appareils à extraire.

  for (let r = hdr.row + 1; r < matrix.length; r++) { // Parcourt les lignes situées après l'en-tête.
    const row = matrix[r] || []; // Récupère la ligne courante ou un tableau vide.
    const get = (f) => (f in hdr.cols ? row[hdr.cols[f]] : null); // Lit une cellule à partir du nom logique de sa colonne.

    const prefix = clean(get('prefix')); // Récupère et nettoie le préfixe de l'appareil.

    if (!prefix) continue; // Ignore les lignes sans préfixe d'appareil.

    const rawNum = get('number'); // Récupère le numéro brut de l'appareil.

    const num = // Prépare le numéro final sous forme de texte.
      typeof rawNum === 'number' // Vérifie si le numéro est stocké comme valeur numérique.
        ? String(Math.trunc(rawNum)).padStart(2, '0') // Supprime les décimales et ajoute un zéro devant les nombres à un chiffre.
        : clean(rawNum); // Nettoie le numéro s'il est déjà stocké comme texte.

    let desc = clean(get('desc')); // Récupère et nettoie la description de l'appareil.

    for (const rm of opts.removeText) { // Parcourt les textes que l'utilisateur souhaite supprimer.
      if (rm) { // Ignore les éléments vides de la liste de suppression.
        desc = desc.replace(new RegExp(escapeRe(rm), 'gi'), ''); // Supprime toutes les occurrences du texte sans tenir compte de la casse.
      } // Termine la vérification du texte à supprimer.
    } // Termine le nettoyage des textes de description.

    desc = desc // Commence la normalisation finale de la description.
      .replace(/\s+/g, ' ') // Remplace les séries d'espaces par un espace unique.
      .trim() // Supprime les espaces au début et à la fin.
      .replace(/^[\s-]+|[\s-]+$/g, ''); // Supprime les espaces et tirets situés aux extrémités.

    devices.push({ // Ajoute un appareil à la liste des appareils extraits.
      srcRow: r + 1, // Conserve le numéro de ligne Excel d'origine, en commençant à 1.
      loop: clean(get('loop')), // Récupère le numéro de boucle indiqué sur la ligne.
      prefix, // Enregistre le préfixe nettoyé de l'appareil.
      num, // Enregistre le numéro nettoyé de l'appareil.
      addr: clean(get('address')), // Récupère l'adresse de l'appareil.
      desc, // Enregistre la description nettoyée.
      area: clean(get('area')), // Récupère la zone de l'appareil.
    }); // Termine la création de l'objet appareil.
  } // Termine la lecture des lignes de la IO LIST.

  return devices; // Retourne la liste des appareils trouvés.
} // Termine la fonction d'analyse de la IO LIST.
// Construit les câbles à partir des appareils extraits de la IO LIST.
function buildRows(devices, opts) { // Reçoit la liste des appareils et les options choisies.
  const warnings = []; // Stocke les avertissements à afficher à l'utilisateur.
  const loops = new Map(); // Stocke les appareils regroupés par numéro de boucle.
  let current = null; // Mémorise la dernière boucle connue pendant la lecture.

  for (const d of devices) { // Parcourt les appareils dans l'ordre de la IO LIST.
    if (d.loop) current = d.loop; // Met à jour la boucle courante si une boucle est indiquée.

    if (current === null) { // Vérifie qu'une boucle a déjà été identifiée.
      warnings.push( // Ajoute un avertissement lorsqu'aucune boucle précédente n'existe.
        `IO LIST row ${d.srcRow} (${d.prefix}-${d.num}): ` + // Indique la ligne source et le repère de l'appareil.
        'no preceding loop found; device skipped' // Explique que l'appareil sera ignoré.
      ); // Termine l'ajout de l'avertissement.
      continue; // Ignore l'appareil tant qu'aucune boucle ne peut lui être attribuée.
    } // Termine la vérification de la boucle courante.

    d.onLoop = Boolean(d.loop && d.addr); // Considère l'appareil comme étant sur la boucle si sa ligne contient une boucle et une adresse.

    if (!loops.has(current)) loops.set(current, []); // Crée la liste de la boucle si elle n'existe pas encore.

    loops.get(current).push(d); // Ajoute l'appareil à la boucle correspondante.
  } // Termine le regroupement des appareils par boucle.

  const rows = []; // Prépare la liste finale des lignes de câbles.
  const hubs = new Set(opts.hubPrefixes); // Transforme les préfixes des hubs en ensemble pour faciliter les recherches.

  const nextRow = () => rows.length + 2; // Calcule le numéro Excel de la prochaine ligne de données, après l'en-tête.

  for (const [loop, devs] of loops) { // Traite chaque boucle et les appareils qui lui appartiennent.
    let prevRow = null; // Mémorise la ligne du dernier appareil raccordé directement à la boucle.
    let hubRow = null; // Mémorise la ligne du hub actif lorsqu'une dérivation passe par un hub.
    let lastArea = devs[0].area; // Initialise la zone de référence avec la zone du premier appareil.

    const tagCount = new Map(); // Compte les occurrences de chaque repère d'appareil dans la boucle.

    for (const d of devs) { // Parcourt les appareils de la boucle dans leur ordre d'origine.
      const tag = `${d.prefix}-${d.num}`; // Construit le repère de l'appareil, par exemple MCP-01.

      tagCount.set(tag, (tagCount.get(tag) || 0) + 1); // Incrémente le nombre d'occurrences de ce repère.

      const r = nextRow(); // Calcule la ligne Excel de sortie de l'appareil courant.
      let srcRow; // Déclare la ligne Excel de l'appareil d'où part le câble.

      if (d.onLoop) { // Traite un appareil raccordé directement à la boucle.
        srcRow = prevRow; // Utilise l'appareil précédent comme origine, ou null pour le premier appareil.
        prevRow = r; // Mémorise l'appareil courant comme dernier appareil de la boucle.
        hubRow = null; // Réinitialise le hub actif après un appareil directement sur la boucle.
      } else { // Traite un appareil raccordé en dérivation.
        const parent = hubRow || prevRow; // Choisit le hub actif, sinon le dernier appareil de la boucle.

        if (parent === null) { // Vérifie qu'aucun appareil précédent ne peut servir d'origine.
          warnings.push( // Ajoute un avertissement concernant la dérivation.
            `${loop} ${tag}: branch connection has no preceding device; ` + // Identifie la boucle et le repère concerné.
            `connected to ${opts.facpName}` // Indique que le câble sera rattaché au FACP.
          ); // Termine l'ajout de l'avertissement.

          srcRow = null; // Utilise null pour représenter le FACP comme origine du câble.
        } else { // Si un appareil précédent est disponible.
          srcRow = parent; // Utilise cet appareil comme origine du câble en dérivation.
        } // Termine le choix de l'origine de la dérivation.

        if (hubs.has(d.prefix)) hubRow = r; // Mémorise le hub courant s'il appartient aux préfixes configurés.
      } // Termine le traitement du raccordement de l'appareil.

      if (d.desc === '') { // Vérifie si la description de l'appareil est absente.
        warnings.push( // Ajoute un avertissement si la description est vide.
          `${loop} ${tag}: missing description in IO LIST` // Identifie l'appareil dont la description manque.
        ); // Termine l'ajout de l'avertissement.
      } // Termine la vérification de la description.

      lastArea = d.area || lastArea; // Met à jour la dernière zone connue si l'appareil possède une zone.

      rows.push({ // Ajoute la ligne représentant le câble de l'appareil.
        tag, // Enregistre le repère de l'appareil.
        loop: d.loop, // Enregistre le numéro de boucle indiqué sur la ligne source.
        kind: 'device', // Indique qu'il s'agit d'une ligne d'appareil et non d'un retour au FACP.
        onLoop: d.onLoop, // Indique si l'appareil est raccordé directement à la boucle.
        srcRow, // Enregistre la ligne Excel de l'origine du câble.
        addr: // Détermine l'adresse à inscrire dans le fichier final.
          opts.addressMode === 'type' // Vérifie si le mode d'adresse par type a été sélectionné.
            ? TYPE_ADDRESS[d.prefix] || '' // Utilise l'adresse définie dans TYPE_ADDRESS, ou une chaîne vide si elle manque.
            : d.addr, // Sinon, utilise l'adresse lue dans la IO LIST.
        desc: d.desc, // Enregistre la description de l'appareil.
        area: d.area, // Enregistre la zone de l'appareil.
      }); // Termine l'ajout de la ligne de câble.
    } // Termine le traitement des appareils de la boucle.

    for (const [t, n] of tagCount) { // Parcourt les repères et leurs nombres d'occurrences.
      if (n > 1) { // Vérifie si un repère apparaît plusieurs fois.
        warnings.push( // Ajoute un avertissement concernant les repères dupliqués.
          `${loop}: tag "${t}" appears ${n} times; duplicate cable labels` // Indique la boucle, le repère et le nombre d'occurrences.
        ); // Termine l'ajout de l'avertissement.
      } // Termine la vérification des doublons.
    } // Termine le contrôle des repères.

    rows.push({ // Ajoute une ligne pour le câble de retour vers le FACP.
      tag: opts.facpName, // Utilise le nom configuré du FACP comme repère de destination.
      loop, // Enregistre le numéro de la boucle concernée.
      kind: 'return', // Identifie cette ligne comme un câble de retour.
      onLoop: true, // Marque le retour comme faisant partie du circuit de boucle.
      srcRow: prevRow, // Utilise le dernier appareil directement sur la boucle comme origine du retour.
      addr: '', // Laisse l'adresse vide pour le câble de retour.
      desc: opts.facpDescription, // Utilise la description configurée du FACP.
      area: lastArea, // Affecte au retour la dernière zone connue de la boucle.
    }); // Termine l'ajout du câble de retour.
  } // Termine le traitement de toutes les boucles.

  // Prépare les repères utilisés dans l'aperçu et dans les résultats de formules Excel.
  for (const row of rows) { // Parcourt toutes les lignes de câbles générées.
    row.label = // Crée le repère complet qui sera affiché dans le fichier final.
      row.loop && row.kind === 'device' // Vérifie qu'il s'agit d'un appareil avec un numéro de boucle.
        ? `${row.loop}-${row.tag}` // Compose le repère à partir de la boucle et de l'appareil.
        : row.tag; // Utilise directement le repère pour les autres lignes.
  } // Termine la création des repères.

  for (const row of rows) { // Parcourt les lignes pour déterminer le libellé de l'origine du câble.
    row.srcLabel = // Stocke le nom de l'origine à afficher dans la colonne FROM.
      row.srcRow === null // Vérifie si l'origine est le FACP plutôt qu'une ligne d'appareil.
        ? opts.facpName // Utilise le nom du FACP comme origine.
        : rows[row.srcRow - 2].label; // Récupère le repère de l'appareil correspondant à la ligne Excel source.
  } // Termine la détermination des origines.

  return { rows, loops, warnings }; // Retourne les câbles, les boucles et les avertissements.
} // Termine la fonction de construction des câbles.
// Crée le classeur Excel final à partir des câbles générés.
function buildWorkbook(rows, opts, ExcelLib) { // Reçoit les lignes de câbles, les options et la bibliothèque ExcelJS.
  const wb = new ExcelLib.Workbook(); // Crée un nouveau classeur Excel.
  const ws = wb.addWorksheet('EQUIPMENT'); // Crée une feuille nommée EQUIPMENT.

  const font = { name: 'Calibri', size: 14 }; // Définit la police standard du tableau.

  const headFont = { // Définit le style de la police des en-têtes.
    name: 'Calibri', // Utilise la police Calibri.
    size: 14, // Définit la taille de la police des en-têtes.
    bold: true, // Met les en-têtes en gras.
    color: { argb: 'FFFFFFFF' }, // Définit la couleur du texte en blanc.
  }; // Termine le style des en-têtes.

  const solid = (argb) => ({ // Déclare une fonction qui crée un remplissage de cellule uni.
    type: 'pattern', // Indique que le remplissage utilise un motif.
    pattern: 'solid', // Définit un remplissage uniforme.
    fgColor: { argb }, // Définit la couleur du remplissage.
  }); // Termine la fonction de création des remplissages.

  const thin = { style: 'thin' }; // Définit une bordure fine.

  const border = { // Regroupe les bordures à appliquer aux cellules.
    top: thin, // Applique une bordure fine en haut.
    left: thin, // Applique une bordure fine à gauche.
    bottom: thin, // Applique une bordure fine en bas.
    right: thin, // Applique une bordure fine à droite.
  }; // Termine la définition des bordures.

  const center = { // Définit l'alignement commun des cellules.
    horizontal: 'center', // Centre le contenu horizontalement.
    vertical: 'middle', // Centre le contenu verticalement.
    wrapText: true, // Autorise le retour à la ligne du texte.
  }; // Termine la définition de l'alignement.

  const layout = [ // Définit les colonnes et les regroupements de cellules de la feuille Excel.
    [1, 1, 'ITEM'], // Utilise la colonne A pour le numéro de ligne.
    [2, 2, 'EQUIPMENT'], // Utilise la colonne B pour le repère de l'appareil.
    [3, 3, 'ADDRESS'], // Utilise la colonne C pour l'adresse.
    [4, 4, 'LOOP No.'], // Utilise la colonne D pour le numéro de boucle.
    [5, 5, 'CABLE LABEL'], // Utilise la colonne E pour le repère du câble.
    [6, 9, 'FROM'], // Fusionne les colonnes F à I pour l'origine du câble.
    [10, 10, 'TO'], // Utilise la colonne J pour la destination du câble.
    [11, 12, 'TYPE'], // Fusionne les colonnes K et L pour le type de câble.
    [13, 13, 'SHIELD'], // Utilise la colonne M pour le blindage.
    [14, 14, 'ARMORED'], // Utilise la colonne N pour l'armure du câble.
    [15, 15, 'LENGTH (m)'], // Utilise la colonne O pour la longueur en mètres.
    [16, 19, 'DESCRIPTION'], // Fusionne les colonnes P à S pour la description.
    [20, 23, 'AREA'], // Fusionne les colonnes T à W pour la zone.
  ]; // Termine la définition des colonnes.

  // Inscrit une valeur, applique le style et fusionne éventuellement les cellules.
  function put(r, c1, c2, value, o = {}) { // Reçoit la ligne, les colonnes de début et de fin, la valeur et les options de style.
    for (let c = c1; c <= c2; c++) { // Parcourt toutes les cellules de la plage à remplir.
      const cell = ws.getCell(r, c); // Récupère la cellule à la ligne et à la colonne courantes.

      cell.border = border; // Applique les bordures définies précédemment.
      cell.alignment = center; // Centre le contenu de la cellule.
      cell.font = o.font || font; // Utilise la police personnalisée ou la police standard.

      if (o.fill) cell.fill = o.fill; // Applique un remplissage si une couleur est fournie.
    } // Termine le formatage des cellules de la plage.

    const master = ws.getCell(r, c1); // Récupère la première cellule de la plage, qui recevra la valeur.

    master.value = // Affecte la valeur à la cellule principale.
      value === '' || value === undefined ? null : value; // Convertit les chaînes vides et les valeurs non définies en cellules vides.

    if (o.numFmt) master.numFmt = o.numFmt; // Applique le format numérique demandé, si nécessaire.

    if (c2 > c1) ws.mergeCells(r, c1, r, c2); // Fusionne les cellules si la plage contient plusieurs colonnes.
  } // Termine la fonction d'écriture des cellules.

  // Crée la ligne des en-têtes du tableau.
  for (const [c1, c2, title] of layout) { // Parcourt chaque groupe de colonnes défini dans layout.
    put(1, c1, c2, title, { // Écrit le titre du groupe sur la première ligne.
      font: headFont, // Applique la police blanche et grasse.
      fill: solid('FF4472C4'), // Applique un fond bleu aux en-têtes.
    }); // Termine l'écriture du titre.
  } // Termine la création des en-têtes.

  ws.getRow(1).height = 40; // Définit la hauteur de la ligne des en-têtes.

  const yellow = solid('FFFFFF00'); // Prépare le remplissage jaune pour les champs à vérifier ou compléter.

  rows.forEach((d, i) => { // Parcourt chaque câble et son indice dans la liste.
    const r = i + 2; // Calcule la ligne Excel du câble, en commençant à la ligne 2.

    const labelFormula = // Prépare la formule Excel utilisée pour afficher la destination.
      d.loop && d.kind === 'device' // Vérifie si la destination est un appareil associé à une boucle.
        ? { // Définit le cas des appareils de boucle.
            formula: `CONCATENATE(D${r},"-",B${r})`, // Construit le repère à partir du numéro de boucle et du repère de l'appareil.
            result: d.label, // Fournit le résultat calculé à l'avance pour l'affichage initial.
          } // Termine le cas des appareils.
        : { // Définit le cas des autres destinations, notamment le FACP.
            formula: `CONCATENATE(B${r})`, // Reprend le contenu de la colonne B comme destination.
            result: d.label, // Fournit le résultat attendu de la formule.
          }; // Termine la préparation de la formule.

    put(r, 1, 1, i + 1, { numFmt: '0' }); // Écrit le numéro séquentiel du câble dans la colonne A.

    put(r, 2, 2, d.tag, { numFmt: '@' }); // Écrit le repère de l'appareil en conservant le format texte.

    put(r, 3, 3, d.addr, { numFmt: '@' }); // Écrit l'adresse en texte pour préserver les zéros initiaux.

    put(r, 4, 4, d.loop, { numFmt: '@' }); // Écrit le numéro de boucle au format texte.

    put(r, 5, 5, { // Écrit le repère du câble dans la colonne E.
      formula: `J${r}`, // Utilise la destination de la colonne J comme formule du repère.
      result: d.label, // Fournit le résultat de la formule pour l'aperçu.
    }); // Termine l'écriture du repère du câble.

    put( // Écrit l'origine du câble dans les colonnes F à I.
      r, // Définit la ligne Excel du câble.
      6, // Définit la première colonne de la plage : F.
      9, // Définit la dernière colonne de la plage : I.
      d.srcRow === null // Vérifie si le câble part directement du FACP.
        ? opts.facpName // Écrit directement le nom du FACP.
        : { // Sinon, crée une référence vers la destination de la ligne source.
            formula: `J${d.srcRow}`, // Récupère le repère de l'origine dans la colonne J de la ligne précédente correspondante.
            result: d.srcLabel, // Fournit le texte de l'origine pour l'aperçu.
          } // Termine la définition de l'origine référencée.
    ); // Termine l'écriture de l'origine.

    put(r, 10, 10, labelFormula); // Écrit la destination du câble dans la colonne J.

    // Les cellules jaunes doivent être vérifiées ou complétées par l'utilisateur.
    put(r, 11, 12, opts.cableType, { fill: yellow }); // Écrit le type de câble dans les colonnes K et L avec un fond jaune.

    put(r, 13, 13, opts.shield); // Écrit la valeur du blindage dans la colonne M.

    put(r, 14, 14, opts.armored); // Écrit la valeur de l'armure dans la colonne N.

    put(r, 15, 15, null, { // Prépare la cellule destinée à la longueur du câble.
      fill: yellow, // Colore la cellule en jaune pour signaler qu'elle doit être complétée.
      numFmt: '0', // Affiche la longueur comme un nombre entier.
    }); // Termine la préparation de la longueur.

    put(r, 16, 19, d.desc); // Écrit la description dans les colonnes P à S.

    put(r, 20, 23, d.area); // Écrit la zone dans les colonnes T à W.

    ws.getRow(r).height = 20.1; // Définit la hauteur de la ligne du câble.
  }); // Termine l'écriture de tous les câbles.

  // Ajoute une ligne noire en bas du tableau.
  const end = rows.length + 2; // Calcule la ligne Excel située après la dernière ligne de câble.

  for (let c = 1; c <= 23; c++) { // Parcourt les 23 colonnes du tableau.
    ws.getCell(end, c).fill = solid('FF000000'); // Applique un remplissage noir à chaque cellule.
  } // Termine la création de la ligne noire.

  ws.mergeCells(end, 1, end, 23); // Fusionne les 23 cellules de la ligne noire.
  ws.getRow(end).height = 6; // Rend la ligne noire très fine.

  const widths = [ // Définit la largeur de chaque colonne Excel.
    10, 16, 12, 10, 26, 7, 7, 7, 7, 26, 8, 8, // Largeurs des colonnes A à L.
    10, 13, 17, 30, 10, 10, 10, 20, 12, 12, 12, // Largeurs des colonnes M à W.
  ]; // Termine la liste des largeurs.

  widths.forEach((w, i) => { // Parcourt les largeurs et leurs indices.
    ws.getColumn(i + 1).width = w; // Affecte la largeur à la colonne correspondante, en commençant à A.
  }); // Termine le réglage des largeurs.

  ws.views = [{ state: 'frozen', ySplit: 1 }]; // Fige la première ligne pour garder les en-têtes visibles pendant le défilement.

  ws.autoFilter = `A1:W${end - 1}`; // Active les filtres Excel sur le tableau, de A1 à W sur la dernière ligne de données.

  ws.pageSetup = { // Configure la mise en page pour l'impression.
    orientation: 'landscape', // Utilise l'orientation paysage.
    fitToPage: true, // Demande l'ajustement du tableau à la page.
    fitToWidth: 1, // Ajuste le tableau à une page en largeur.
    fitToHeight: 0, // Laisse Excel utiliser autant de pages que nécessaire en hauteur.
  }; // Termine la configuration d'impression.

  return wb; // Retourne le classeur prêt à être exporté.
} // Termine la génération du classeur Excel.
/* Export pour Node.js : permet d'utiliser la logique sans lancer l'interface du navigateur. */

if (typeof module !== 'undefined' && module.exports) { // Vérifie que le programme est exécuté dans un environnement Node.js compatible CommonJS.
  module.exports = { // Expose les fonctions et constantes pour les tests ou les autres modules.
    DEFAULTS, // Exporte les paramètres par défaut.
    parseIoMatrix, // Exporte la fonction d'analyse de la IO LIST.
    buildRows, // Exporte la fonction de construction des lignes de câbles.
    buildWorkbook, // Exporte la fonction de création du classeur Excel.
    findHeader, // Exporte la fonction de recherche des en-têtes.
    cellText, // Exporte la fonction de conversion des cellules Excel.
  }; // Termine la liste des éléments exportés.
} // Termine la vérification de l'environnement Node.js.
// =============================================================================
// 4. INTERFACE UTILISATEUR : ce bloc s'exécute uniquement dans un navigateur.
// =============================================================================

if (typeof document !== 'undefined') { // Vérifie que l'objet document existe avant d'utiliser le DOM.
  const $ = (id) => document.getElementById(id); // Simplifie la recherche d'un élément HTML à partir de son identifiant.

  const state = { // Centralise les informations conservées pendant l'utilisation de l'application.
    fileName: '', // Stocke le nom du fichier Excel chargé.
    candidates: [], // Stocke les feuilles Excel contenant les colonnes requises.
    result: null, // Stocke le résultat de la génération des câbles.
  }; // Termine la définition de l'état.

  const el = { // Regroupe les références aux éléments HTML de l'interface.
    drop: $('drop'), // Référence à la zone de glisser-déposer du fichier.
    file: $('file'), // Référence au champ de sélection du fichier.
    fileInfo: $('file-info'), // Référence à l'élément affichant le nom du fichier.
    sheetRow: $('sheet-row'), // Référence à la ligne HTML contenant le sélecteur de feuille.
    sheet: $('sheet'), // Référence au sélecteur de feuille Excel.
    status: $('status'), // Référence à la zone affichant les messages d'état.
    results: $('results'), // Référence au conteneur des résultats.
    empty: $('empty'), // Référence au message affiché lorsqu'aucun résultat n'est disponible.
    stats: $('stats'), // Référence à la zone des statistiques.
    loopBody: $('loop-body'), // Référence au corps du tableau récapitulatif des boucles.
    warnBox: $('warn-box'), // Référence au conteneur des avertissements.
    warnList: $('warn-list'), // Référence à la liste des avertissements.
    warnCount: $('warn-count'), // Référence au compteur d'avertissements.
    preview: $('preview-body'), // Référence au corps du tableau d'aperçu des câbles.
    download: $('download'), // Référence au bouton de téléchargement.
    addressMode: $('opt-address'), // Référence au choix du mode d'adresse.
    cableType: $('opt-cable'), // Référence au champ du type de câble.
    facp: $('opt-facp'), // Référence au champ du nom du FACP.
    remove: $('opt-remove'), // Référence au champ des textes à supprimer.
    hubs: $('opt-hubs'), // Référence au champ des préfixes des hubs.
    reset: $('opt-reset'), // Référence au bouton de réinitialisation des options.
  }; // Termine le regroupement des éléments HTML.

// Affiche un message dans la zone d'état de l'interface.
  function setStatus(msg, kind) { // Reçoit le texte du message et sa catégorie.
    el.status.textContent = msg || ''; // Affiche le message ou une chaîne vide.
    el.status.dataset.kind = kind || ''; // Enregistre la catégorie pour permettre un style CSS adapté.
  } // Termine la fonction d'affichage des messages.

// Lit les paramètres actuels saisis dans l'interface.
  function readOptions() { // Construit un objet d'options à partir des champs HTML.
    const split = (s) => // Déclare une fonction qui transforme une liste séparée par des virgules en tableau.
      s.split(',').map((x) => x.trim()).filter(Boolean); // Sépare les éléments, supprime les espaces et élimine les entrées vides.

    return { // Retourne toutes les options applicables à la génération.
      ...DEFAULTS, // Commence par copier les valeurs par défaut.
      addressMode: el.addressMode.value, // Récupère le mode d'adresse sélectionné.
      cableType: // Détermine le type de câble à utiliser.
        el.cableType.value.trim() || DEFAULTS.cableType, // Utilise la valeur saisie ou le type par défaut.
      facpName: // Détermine le nom du panneau FACP.
        el.facp.value.trim() || DEFAULTS.facpName, // Utilise le nom saisi ou le nom par défaut.
      removeText: split(el.remove.value), // Transforme les textes à supprimer en tableau.
      hubPrefixes: split(el.hubs.value.toUpperCase()), // Transforme les préfixes des hubs en tableau de majuscules.
    }; // Termine la construction des options.
  } // Termine la lecture des options.

// Réinitialise les champs de l'interface à leurs valeurs initiales.
  function resetOptions() { // Déclare la fonction de réinitialisation.
    el.addressMode.value = DEFAULTS.addressMode; // Restaure le mode d'adresse par défaut.
    el.cableType.value = DEFAULTS.cableType; // Restaure le type de câble par défaut.
    el.facp.value = DEFAULTS.facpName; // Restaure le nom du FACP.
    el.remove.value = DEFAULTS.removeText.join(', '); // Restaure la liste des textes à supprimer.
    el.hubs.value = DEFAULTS.hubPrefixes.join(', '); // Restaure la liste des préfixes des hubs.
  } // Termine la réinitialisation des options.

// Convertit une feuille ExcelJS en matrice de valeurs simples.
  function sheetToMatrix(ws) { // Reçoit une feuille ExcelJS.
    const matrix = []; // Prépare la matrice qui contiendra les lignes.
    const cols = ws.columnCount; // Récupère le nombre de colonnes de la feuille.

    for (let r = 1; r <= ws.rowCount; r++) { // Parcourt toutes les lignes de la feuille Excel.
      const row = ws.getRow(r); // Récupère la ligne Excel courante.
      const arr = []; // Prépare le tableau de valeurs de cette ligne.

      for (let c = 1; c <= cols; c++) { // Parcourt toutes les colonnes de la ligne.
        arr.push(cellText(row.getCell(c).value)); // Convertit chaque cellule en valeur simple et l'ajoute à la ligne.
      } // Termine la lecture des colonnes.

      matrix.push(arr); // Ajoute la ligne à la matrice.
    } // Termine la lecture des lignes.

    return matrix; // Retourne la matrice complète de la feuille.
  } // Termine la conversion de la feuille.

// Charge et analyse un fichier Excel fourni par l'utilisateur.
  async function loadFile(file) { // Reçoit le fichier sélectionné ou déposé.
    if (!file) return; // Arrête la fonction si aucun fichier n'a été fourni.

    if (typeof ExcelJS === 'undefined') { // Vérifie que la bibliothèque ExcelJS est disponible.
      setStatus( // Affiche une erreur si la bibliothèque n'a pas été chargée.
        'ExcelJS library could not be loaded. Check your internet connection and reload the page.', // Explique que le chargement d'ExcelJS a échoué.
        'error' // Marque le message comme une erreur.
      ); // Termine l'affichage du message.
      return; // Interrompt le traitement du fichier.
    } // Termine la vérification d'ExcelJS.

    if (!/\.xlsx$/i.test(file.name)) { // Vérifie que le fichier porte l'extension .xlsx.
      setStatus( // Affiche une erreur si le format n'est pas accepté.
        'Unsupported format: upload a .xlsx file. ' + // Indique que seul le format XLSX est accepté.
        'If your IO list is in .xls format, save it as .xlsx first.', // Explique comment préparer un ancien fichier XLS.
        'error' // Marque le message comme une erreur.
      ); // Termine l'affichage du message.
      return; // Arrête le traitement du fichier non pris en charge.
    } // Termine la vérification du format.

    setStatus('Reading file…'); // Informe l'utilisateur que la lecture du fichier commence.

    try { // Commence un bloc permettant de gérer les erreurs de lecture.
      const wb = new ExcelJS.Workbook(); // Crée un objet classeur ExcelJS.
      await wb.xlsx.load(await file.arrayBuffer()); // Lit le fichier en mémoire puis charge son contenu dans ExcelJS.

      state.fileName = file.name; // Mémorise le nom du fichier sélectionné.
      state.candidates = []; // Réinitialise la liste des feuilles candidates.

      wb.eachSheet((ws) => { // Parcourt toutes les feuilles du classeur.
        const matrix = sheetToMatrix(ws); // Convertit la feuille courante en matrice de valeurs.

        if (findHeader(matrix)) { // Vérifie si la feuille contient les colonnes obligatoires.
          state.candidates.push({ // Ajoute la feuille à la liste des feuilles compatibles.
            name: ws.name, // Stocke le nom de la feuille.
            matrix, // Stocke son contenu sous forme de matrice.
          }); // Termine l'ajout de la feuille candidate.
        } // Termine la vérification de la feuille.
      }); // Termine le parcours du classeur.

      el.fileInfo.textContent = file.name; // Affiche le nom du fichier chargé.

      if (!state.candidates.length) { // Vérifie si aucune feuille compatible n'a été trouvée.
        el.results.hidden = true; // Masque les résultats précédents.
        el.empty.hidden = false; // Affiche le message indiquant l'absence de données compatibles.
        el.download.disabled = true; // Désactive le téléchargement.

        setStatus( // Affiche un message expliquant le problème de structure.
          'No sheet contains the columns LOOP, PREFISSO and ' + // Indique les premières colonnes obligatoires.
          'NUMERO DELL’APPARECCHIATURA. Check the IO list header row.', // Demande de vérifier la ligne des en-têtes.
          'error' // Marque le message comme une erreur.
        ); // Termine l'affichage du message.

        return; // Arrête la fonction puisqu'aucune feuille exploitable n'existe.
      } // Termine le contrôle des feuilles candidates.

      el.sheet.innerHTML = ''; // Vide les anciennes options du sélecteur de feuilles.

      state.candidates.forEach((c, i) => { // Parcourt les feuilles compatibles.
        el.sheet.add(new Option(c.name, i)); // Ajoute le nom de chaque feuille au sélecteur.
      }); // Termine le remplissage du sélecteur.

      el.sheetRow.hidden = state.candidates.length < 2; // Masque le sélecteur s'il n'existe qu'une seule feuille compatible.

      generate(); // Lance automatiquement la génération à partir de la feuille sélectionnée.
    } catch (err) { // Intercepte les erreurs pouvant survenir pendant la lecture.
      console.error(err); // Enregistre les détails techniques de l'erreur dans la console.

      setStatus( // Affiche un message compréhensible pour l'utilisateur.
        'Unable to read this file. It may be password-protected or corrupted.', // Indique que le fichier est peut-être protégé ou endommagé.
        'error' // Marque le message comme une erreur.
      ); // Termine l'affichage du message.
    } // Termine la gestion des erreurs de lecture.
  } // Termine le chargement du fichier.

// Génère la liste de câbles à partir de la feuille et des options sélectionnées.
  function generate() { // Déclare la fonction principale de génération.
    const cand = state.candidates[Number(el.sheet.value) || 0]; // Récupère la feuille choisie, ou la première feuille par défaut.

    if (!cand) return; // Arrête le traitement si aucune feuille n'est sélectionnée.

    const opts = readOptions(); // Lit les options actuelles de l'utilisateur.
    const devices = parseIoMatrix(cand.matrix, opts); // Extrait les appareils de la feuille sélectionnée.
    const res = buildRows(devices, opts); // Construit les lignes de câbles et les avertissements.

    state.result = { // Mémorise le résultat de la génération.
      ...res, // Copie les lignes, les boucles et les avertissements.
      opts, // Conserve les options utilisées.
      sheet: cand.name, // Conserve le nom de la feuille traitée.
      deviceCount: devices.length, // Enregistre le nombre d'appareils extraits.
    }; // Termine l'enregistrement du résultat.

    render(); // Actualise les statistiques, les avertissements et l'aperçu.

    setStatus( // Affiche un message de confirmation de la génération.
      `Cable list ready: ${res.rows.length} cables across ${res.loops.size} loops.`, // Indique le nombre de câbles et de boucles.
      'ok' // Marque le message comme un succès.
    ); // Termine l'affichage du message.
  } // Termine la génération.

// Crée une cellule HTML de tableau et lui affecte du texte.
  const td = (text, cls) => { // Déclare une fonction qui simplifie la création de cellules td.
    const c = document.createElement('td'); // Crée une cellule de tableau HTML.

    c.textContent = text == null ? '' : text; // Inscrit le texte sans interpréter son contenu comme du HTML.

    if (cls) c.className = cls; // Applique une classe CSS si elle a été fournie.

    return c; // Retourne la cellule créée.
  }; // Termine la fonction de création de cellule.

// Affiche les résultats calculés dans l'interface.
  function render() { // Déclare la fonction d'affichage.
    const { rows, loops, warnings, deviceCount } = state.result; // Extrait les données nécessaires depuis le résultat.

    el.empty.hidden = true; // Masque le message indiquant l'absence de résultats.
    el.results.hidden = false; // Affiche la zone contenant les résultats.
    el.download.disabled = false; // Active le bouton de téléchargement.

    // Affiche les statistiques générales.
    el.stats.innerHTML = ''; // Efface les anciennes statistiques.

    [ // Prépare les trois statistiques à afficher.
      [deviceCount, 'Devices read'], // Indique le nombre d'appareils lus.
      [loops.size, 'Loops'], // Indique le nombre de boucles identifiées.
      [rows.length, 'Cables generated'], // Indique le nombre de lignes de câbles générées.
    ].forEach(([n, label]) => { // Parcourt chaque statistique et son libellé.
      const d = document.createElement('div'); // Crée un conteneur HTML pour la statistique.

      const b = document.createElement('strong'); // Crée un élément destiné à afficher le nombre.
      b.textContent = n; // Inscrit la valeur numérique de la statistique.

      const s = document.createElement('span'); // Crée un élément destiné au libellé.
      s.textContent = label; // Inscrit le texte explicatif de la statistique.

      d.append(b, s); // Place le nombre et son libellé dans le conteneur.
      el.stats.append(d); // Ajoute la statistique à la zone des statistiques.
    }); // Termine l'affichage des statistiques.

    // Affiche le récapitulatif des boucles.
    el.loopBody.innerHTML = ''; // Efface les anciennes lignes du récapitulatif.

    for (const [loop, devs] of loops) { // Parcourt les boucles et leurs appareils.
      const onLoop = devs.filter((d) => d.onLoop).length; // Compte les appareils raccordés directement à la boucle.

      const tr = document.createElement('tr'); // Crée une nouvelle ligne HTML pour la boucle.

      tr.append( // Ajoute les cellules de la ligne récapitulative.
        td(loop, 'k'), // Affiche le numéro de boucle.
        td(onLoop), // Affiche le nombre d'appareils directement sur la boucle.
        td(devs.length - onLoop), // Affiche le nombre d'appareils en dérivation.
        td(devs.length + 1) // Affiche le nombre de câbles, en ajoutant le retour vers le FACP.
      ); // Termine l'ajout des cellules.

      el.loopBody.append(tr); // Ajoute la ligne au tableau récapitulatif.
    } // Termine le récapitulatif des boucles.

    // Affiche les avertissements éventuels.
    el.warnBox.hidden = warnings.length === 0; // Masque la zone d'avertissements lorsqu'il n'y en a aucun.

    el.warnCount.textContent = // Prépare le texte du compteur d'avertissements.
      warnings.length === 1 // Vérifie s'il existe un seul avertissement.
        ? '1 item to check' // Utilise le singulier.
        : `${warnings.length} items to check`; // Utilise le pluriel et affiche le nombre.

    el.warnList.innerHTML = ''; // Efface les anciens avertissements.

    warnings.forEach((w) => { // Parcourt les avertissements générés.
      const li = document.createElement('li'); // Crée un élément de liste HTML.
      li.textContent = w; // Inscrit le texte de l'avertissement.
      el.warnList.append(li); // Ajoute l'avertissement à la liste.
    }); // Termine l'affichage des avertissements.

    // Affiche l'aperçu des câbles générés.
    el.preview.innerHTML = ''; // Efface l'ancien aperçu.
    const frag = document.createDocumentFragment(); // Crée un fragment pour regrouper les lignes avant leur insertion.

    rows.forEach((d, i) => { // Parcourt les câbles et leurs indices.
      const tr = document.createElement('tr'); // Crée une nouvelle ligne HTML.

      if (d.kind === 'return') { // Vérifie si la ligne représente le retour vers le FACP.
        tr.className = 'ret'; // Applique la classe CSS des câbles de retour.
      } else if (!d.onLoop) { // Vérifie si l'appareil est raccordé en dérivation.
        tr.className = 'spur'; // Applique la classe CSS des dérivations.
      } // Termine l'attribution des classes.

      tr.append( // Ajoute les cellules de l'aperçu.
        td(i + 1, 'n'), // Affiche le numéro de ligne.
        td(d.label, 'k'), // Affiche le repère complet du câble.
        td(d.addr), // Affiche l'adresse.
        td(d.srcLabel), // Affiche l'origine du câble.
        td(d.desc), // Affiche la description de l'appareil.
        td(d.area) // Affiche la zone.
      ); // Termine l'ajout des cellules.

      frag.append(tr); // Ajoute la ligne au fragment temporaire.
    }); // Termine la construction de l'aperçu.

    el.preview.append(frag); // Insère toutes les lignes dans le tableau d'aperçu.
  } // Termine l'affichage des résultats.

// Génère et télécharge le classeur Excel.
  async function download() { // Déclare une fonction asynchrone de téléchargement.
    if (!state.result) return; // Arrête la fonction si aucun résultat n'a été généré.

    try { // Commence le bloc de gestion des erreurs de génération.
      const wb = buildWorkbook( // Construit le classeur final.
        state.result.rows, // Transmet les lignes de câbles.
        state.result.opts, // Transmet les options utilisées.
        ExcelJS // Transmet la bibliothèque ExcelJS.
      ); // Termine la construction du classeur.

      const buf = await wb.xlsx.writeBuffer(); // Convertit le classeur en données binaires XLSX.

      const blob = new Blob([buf], { // Crée un objet Blob à partir des données du classeur.
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', // Définit le type MIME du fichier Excel XLSX.
      }); // Termine la création du Blob.

      const url = URL.createObjectURL(blob); // Crée une URL temporaire permettant de télécharger le fichier.
      const a = document.createElement('a'); // Crée un lien HTML pour lancer le téléchargement.

      a.href = url; // Affecte l'URL temporaire au lien.
      a.download = 'CABLE_LIST_generated.xlsx'; // Définit le nom du fichier téléchargé.

      document.body.append(a); // Ajoute temporairement le lien au document.
      a.click(); // Déclenche automatiquement le téléchargement.
      a.remove(); // Supprime le lien temporaire du document.

      setTimeout(() => URL.revokeObjectURL(url), 2000); // Libère l'URL temporaire après deux secondes.

      setStatus('CABLE_LIST_generated.xlsx downloaded.', 'ok'); // Confirme que le téléchargement a été lancé.
    } catch (err) { // Intercepte les erreurs de création ou de téléchargement.
      console.error(err); // Affiche les détails techniques dans la console.

      setStatus( // Informe l'utilisateur de l'échec.
        'Failed to generate the Excel file. Reload the page and try again.', // Propose de recharger la page et de réessayer.
        'error' // Marque le message comme une erreur.
      ); // Termine l'affichage du message.
    } // Termine la gestion des erreurs.
  } // Termine la fonction de téléchargement.

// =============================================================================
// ÉCOUTEURS D'ÉVÉNEMENTS : relient les actions de l'utilisateur aux fonctions.
// =============================================================================

// Charge le fichier lorsque l'utilisateur en sélectionne un.
  el.file.addEventListener('change', (e) => { // Écoute le changement du champ de sélection du fichier.
    loadFile(e.target.files[0]); // Transmet le premier fichier sélectionné à la fonction de chargement.
  }); // Termine l'écoute de la sélection du fichier.

// Gère les événements de glisser-déposer lorsque le fichier survole la zone.
  ['dragenter', 'dragover'].forEach((t) => { // Parcourt les deux événements de survol.
    el.drop.addEventListener(t, (e) => { // Ajoute un gestionnaire pour chaque événement.
      e.preventDefault(); // Empêche le navigateur d'utiliser son comportement par défaut.
      el.drop.classList.add('over'); // Ajoute la classe CSS indiquant que la zone est survolée.
    }); // Termine le gestionnaire de survol.
  }); // Termine l'enregistrement des événements de survol.

// Nettoie l'état visuel lorsque le fichier quitte la zone ou est déposé.
  ['dragleave', 'drop'].forEach((t) => { // Parcourt les événements de sortie et de dépôt.
    el.drop.addEventListener(t, (e) => { // Ajoute un gestionnaire pour chaque événement.
      e.preventDefault(); // Empêche le comportement par défaut du navigateur.
      el.drop.classList.remove('over'); // Retire la classe CSS de survol.
    }); // Termine le gestionnaire.
  }); // Termine l'enregistrement des événements de sortie et de dépôt.

// Charge le fichier lorsqu'il est déposé dans la zone prévue.
  el.drop.addEventListener('drop', (e) => { // Écoute le dépôt d'un fichier.
    loadFile(e.dataTransfer.files[0]); // Récupère le premier fichier déposé et le charge.
  }); // Termine l'écoute du dépôt.

// Régénère les résultats lorsque l'utilisateur change de feuille Excel.
  el.sheet.addEventListener('change', generate); // Relance la génération avec la feuille sélectionnée.

// Régénère les câbles lorsque l'utilisateur modifie une option.
  [ // Prépare la liste des champs dont la modification doit actualiser les résultats.
    el.addressMode, // Ajoute le champ du mode d'adresse.
    el.cableType, // Ajoute le champ du type de câble.
    el.facp, // Ajoute le champ du nom du FACP.
    el.remove, // Ajoute le champ des textes à supprimer.
    el.hubs, // Ajoute le champ des préfixes des hubs.
  ].forEach((i) => { // Parcourt chaque champ de configuration.
    i.addEventListener('change', () => { // Écoute la modification de ce champ.
      if (state.candidates.length) generate(); // Relance la génération si des feuilles compatibles ont été chargées.
    }); // Termine l'écoute du champ.
  }); // Termine l'enregistrement des champs de configuration.

// Réinitialise les options lorsque l'utilisateur clique sur le bouton correspondant.
  el.reset.addEventListener('click', () => { // Écoute le clic sur le bouton de réinitialisation.
    resetOptions(); // Restaure les valeurs par défaut dans l'interface.

    if (state.candidates.length) generate(); // Actualise les résultats si un fichier a déjà été chargé.
  }); // Termine l'écoute du bouton de réinitialisation.

// Lance le téléchargement du fichier lorsque l'utilisateur clique sur le bouton.
  el.download.addEventListener('click', download); // Associe le bouton à la fonction de téléchargement.

  resetOptions(); // Initialise les champs de configuration lors du chargement de l'interface.
} // Termine le bloc réservé au navigateur.
