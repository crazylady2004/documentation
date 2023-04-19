#install.packages("readxl")
#install.packages("utils")
#install.packages("clipr")


Debut <- "<!-- wp:epfl/table-filter {\"largeDisplay\":true,\"tableHeaderOptions\":\"header,sort\"} --><!-- wp:table --><figure class=\"wp-block-table\"><table><tbody>"
Fin <- "</tbody></table></figure><!-- /wp:table --><!-- /wp:epfl/table-filter -->"

file_path <- file.choose() # Ouvre une fenêtre de sélection de fichier
if (is.na(file_path)) {  # Vérifie si un fichier a été sélectionné
  cat("Aucun fichier sélectionné.\n")

} else {
  
  file <- read_excel(file_path) # Ouvre le fichier sélectionné avec readxl
  
  # obtenir le nombre de colonnes utilisées dans la première feuille
  DerniereColonne <- ncol(file)
  
  # obtenir le nombre de lignes utilisées dans la première feuille
  DerniereLigne <- nrow(file)
  
  
  # créer le contenu de l'en-tête
  Header <- ""
  for (i in 1:DerniereColonne) {
    headerText <- file[1,i] # obtenir le texte de l'en-tête à partir de la première ligne de chaque colonne
    Header <- paste0(Header, "<td>", headerText, "</td>")
  }
  Header <- paste0("<tr>", Header, "</tr>")
  
  # créer le contenu du corps de la table
  body <- ""
  for (i in 2:DerniereLigne) {
    ligne <- ""
    for (l in 1:DerniereColonne) {
      contenu_cell <- file[i,l]
      contenu_cell <- gsub(">", "&gt;", contenu_cell) # remplacement de > par du html
      contenu_cell <- gsub("<", "&lt;", contenu_cell) # remplacement de < par du html
      ligne <- paste0(ligne, "<td>", contenu_cell, "</td>")
    }
    body <- paste0(body, "<tr>", ligne, "</tr>")
  }
  
  result_html <- paste0(Debut, Header, body, Fin) # concaténation de tout les informations
  clipr::write_clip(result_html) # écrire le code HTML dans le presse-papiers
  cat("Votre texte est copié dans le presse-papiers\n") # message de confirmation du bon fonctionnement de la copie}
}