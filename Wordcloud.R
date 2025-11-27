install.packages(c("tm", "wordcloud","SnowballC"))

# Load libraries
library(tm)
library(wordcloud)
library(SnowballC)


# Import CSV File

df <- read.csv2("C:/Users/gmatejka/Dropbox/08 - Offenerhaushalt/offenerhaushalt_30301_2025_va_fhh.csv")




# Create Word-Cloud

# Pakete laden
library(tm)
library(wordcloud)
library(RColorBrewer)
# Pakete laden
library(tm)
library(wordcloud)
library(RColorBrewer)
library(openxlsx)  # Falls nicht installiert: install.packages("openxlsx")

# Ordnerpfad definieren
folder_path <- "C:/Users/gmatejka/Dropbox/08 - Offenerhaushalt/"
output_path <- "C:/Users/gmatejka/Dropbox/08 - Offenerhaushalt/wordclouds/"

# Ausgabeordner erstellen, falls nicht vorhanden
if (!dir.exists(output_path)) {
  dir.create(output_path)
}

# Alle CSV-Dateien im Ordner finden
csv_files <- list.files(path = folder_path, pattern = "\\.csv$", full.names = TRUE)

# Leerer Dataframe für die Zusammenfassung
summary_df <- data.frame(
  Gemeindename = character(),
  Haushalt = character(),
  Wort_1 = character(),
  Wort_2 = character(),
  Wort_3 = character(),
  Wort_4 = character(),
  Wort_5 = character(),
  stringsAsFactors = FALSE
)

# Funktion zur Wordcloud-Erstellung
create_wordcloud <- function(csv_file, output_path) {
  
  # CSV einlesen
  df <- read.csv2(csv_file)
  
  # Prüfen ob benötigte Spalten existieren
  required_cols <- c("Konto.Text", "Gemeindename", "Haushalt")
  missing_cols <- required_cols[!required_cols %in% colnames(df)]
  
  if (length(missing_cols) > 0) {
    message(paste("Überspringe:", basename(csv_file), "- Fehlende Spalten:", paste(missing_cols, collapse = ", ")))
    return(NULL)
  }
  
  # Gemeindename und Haushalt extrahieren (erster Wert)
  gemeinde <- unique(df$Gemeindename)[1]
  haushalt <- unique(df$Haushalt)[1]
  
  # Sonderzeichen für Dateinamen entfernen
  safe_gemeinde <- gsub("[^[:alnum:]]", "_", gemeinde)
  safe_haushalt <- gsub("[^[:alnum:]]", "_", haushalt)
  
  # Corpus erstellen
  mooncloud <- Corpus(VectorSource(df$Konto.Text))
  
  # Textbereinigung
  mooncloud <- tm_map(mooncloud, content_transformer(tolower))
  mooncloud <- tm_map(mooncloud, removePunctuation)
  mooncloud <- tm_map(mooncloud, removeNumbers)
  mooncloud <- tm_map(mooncloud, removeWords, stopwords("german"))
  mooncloud <- tm_map(mooncloud, stripWhitespace)
  
  # Eigene Stoppwörter (optional)
  eigene_stopwords <- c("diverse", "sonstige", "übrige", "allgemeine", "bzw")
  mooncloud <- tm_map(mooncloud, removeWords, eigene_stopwords)
  
  # Term-Document-Matrix erstellen
  tdm <- TermDocumentMatrix(mooncloud)
  m <- as.matrix(tdm)
  word_freqs <- sort(rowSums(m), decreasing = TRUE)
  
  # Top 5 Wörter extrahieren
  top_5_words <- head(names(word_freqs), 5)
  # Falls weniger als 5 Wörter, mit NA auffüllen
  while (length(top_5_words) < 5) {
    top_5_words <- c(top_5_words, NA)
  }
  
  # PNG-Datei erstellen
  png_file <- paste0(output_path, safe_gemeinde, "_", safe_haushalt, "_wordcloud.png")
  png(png_file, width = 800, height = 800, res = 150)
  
  set.seed(42)
  wordcloud(
    words = names(word_freqs),
    freq = word_freqs,
    min.freq = 2,
    max.words = 100,
    random.order = FALSE,
    colors = brewer.pal(8, "Dark2"),
    scale = c(3, 0.5)
  )
  
  dev.off()
  
  message(paste("Erstellt:", basename(png_file)))
  
  # Ergebnis zurückgeben
  return(data.frame(
    Gemeindename = gemeinde,
    Haushalt = haushalt,
    Wort_1 = top_5_words[1],
    Wort_2 = top_5_words[2],
    Wort_3 = top_5_words[3],
    Wort_4 = top_5_words[4],
    Wort_5 = top_5_words[5],
    stringsAsFactors = FALSE
  ))
}

# Alle CSVs verarbeiten
for (csv_file in csv_files) {
  tryCatch({
    result <- create_wordcloud(csv_file, output_path)
    if (!is.null(result)) {
      summary_df <- rbind(summary_df, result)
    }
  }, error = function(e) {
    message(paste("Fehler bei", basename(csv_file), ":", e$message))
  })
}

# Excel-Datei erstellen
excel_file <- paste0(output_path, "Wordcloud_Zusammenfassung.xlsx")
write.xlsx(summary_df, excel_file, rowNames = FALSE)

message(paste("\n========================================"))
message(paste("Fertig!", nrow(summary_df), "Wordclouds erstellt."))
message(paste("Excel-Zusammenfassung:", excel_file))
message(paste("========================================"))

# Vorschau der Zusammenfassung
print(summary_df)




# Pakete laden
library(dplyr)
library(openxlsx)

# Ordnerpfad definieren
folder_path <- "C:/Users/gmatejka/Dropbox/08 - Offenerhaushalt/"

# Alle CSV-Dateien einlesen und zusammenführen
csv_files <- list.files(path = folder_path, pattern = "\\.csv$", full.names = TRUE)

all_data <- do.call(rbind, lapply(csv_files, function(file) {
  tryCatch({
    read.csv2(file)
  }, error = function(e) {
    message(paste("Fehler beim Lesen:", basename(file)))
    return(NULL)
  })
}))

# Mehrere Suchbegriffe
suchbegriffe <- c("versicherungen", "personal", "energie", "miete", "zinsen")

# 1. Gesamtanzahl der Einträge pro Gemeinde/Haushalt
gesamt_eintraege <- all_data %>%
  group_by(Gemeindename, Haushalt) %>%
  summarise(
    Gesamt_Eintraege = n(),
    Gesamt_Wert = sum(Wert, na.rm = TRUE),
    .groups = "drop"
  )

# 2. Gesamtwert pro MVAG-Gruppe pro Gemeinde/Haushalt
mvag_gesamt <- all_data %>%
  group_by(Gemeindename, Haushalt, Mvag) %>%
  summarise(
    Mvag_Gesamt_Wert = sum(Wert, na.rm = TRUE),
    Mvag_Gesamt_Eintraege = n(),
    .groups = "drop"
  )

# Funktion für Analyse eines Suchbegriffs
analyse_suchbegriff <- function(data, begriff, gesamt, mvag) {
  data %>%
    filter(grepl(begriff, Konto.Text, ignore.case = TRUE)) %>%
    group_by(Gemeindename, Haushalt, Mvag) %>%
    summarise(
      Anzahl_Treffer = n(),
      Summe_Wert = sum(Wert, na.rm = TRUE),
      .groups = "drop"
    ) %>%
    left_join(gesamt, by = c("Gemeindename", "Haushalt")) %>%
    left_join(mvag, by = c("Gemeindename", "Haushalt", "Mvag")) %>%
    mutate(
      Suchbegriff = begriff,
      Anteil_Eintraege_Prozent = round(Anzahl_Treffer / Gesamt_Eintraege * 100, 2),
      Anteil_Wert_Gesamt_Prozent = round(Summe_Wert / Gesamt_Wert * 100, 2),
      Anteil_Wert_Mvag_Prozent = round(Summe_Wert / Mvag_Gesamt_Wert * 100, 2)
    ) %>%
    select(
      Suchbegriff,
      Gemeindename,
      Haushalt,
      Mvag,
      Anzahl_Treffer,
      Gesamt_Eintraege,
      Anteil_Eintraege_Prozent,
      Summe_Wert,
      Gesamt_Wert,
      Anteil_Wert_Gesamt_Prozent,
      Mvag_Gesamt_Wert,
      Anteil_Wert_Mvag_Prozent
    )
}

# Alle Suchbegriffe analysieren
alle_analysen <- do.call(rbind, lapply(suchbegriffe, function(b) {
  analyse_suchbegriff(all_data, b, gesamt_eintraege, mvag_gesamt)
}))

# Sortieren
alle_analysen <- alle_analysen %>%
  arrange(Suchbegriff, Gemeindename, Mvag)

# Ergebnis anzeigen
print(alle_analysen)

# Als Excel mit separaten Sheets speichern
wb <- createWorkbook()

addWorksheet(wb, "Gesamtübersicht")
writeData(wb, "Gesamtübersicht", alle_analysen)

for (begriff in suchbegriffe) {
  sheet_data <- alle_analysen %>% filter(Suchbegriff == begriff)
  addWorksheet(wb, begriff)
  writeData(wb, begriff, sheet_data)
}

saveWorkbook(wb, paste0(folder_path, "Analyse_Suchbegriffe_mit_Mvag.xlsx"), overwrite = TRUE)

message("Excel-Datei erstellt: Analyse_Suchbegriffe_mit_Mvag.xlsx")
