# --- FİNANSAL ZEKA VE ML PANELİ ---

# Gerekli Paketler
library(shiny)
library(shinydashboard)
library(shinyWidgets) # Arama çubuklu menü için eklendi
library(readxl)       # Excel okumak için eklendi
library(dplyr)
library(tidyr)
library(janitor)
library(ggplot2)
library(plotly)
library(DT) 

# --- 1. UI (KULLANICI ARAYÜZÜ) ---
ui <- dashboardPage(
  skin = "blue",
  
  dashboardHeader(title = "Finansal Zeka Paneli", titleWidth = 300),
  
  dashboardSidebar(
    width = 300,
    sidebarMenu(id = "ana_menu", 
                menuItem("Sektör Görünümü", tabName = "sektor", icon = icon("chart-line")),
                menuItem("Kurum Analizi", tabName = "kurum", icon = icon("building")),
                menuItem("Makine Öğrenmesi", tabName = "ml", icon = icon("robot"))
    ),
    
    hr(),
    h4("  Veri Yönetimi", style = "color: white; padding-left: 15px;"),
    # YENİ: Excel Yükleme Butonu
    fileInput(inputId = "yeni_excel", 
              label = "Yeni Çeyrek Verisi (Excel)",
              accept = c(".xlsx", ".xls"), 
              buttonLabel = "Yükle...",
              placeholder = "Veri seçilmedi"),
    
    hr(),
    h4("  Genel Filtreler", style = "color: white; padding-left: 15px;"),
    # Başlangıçta boş kalıp Server tarafından doldurulacaklar
    selectInput(inputId = "secilen_yil", label = "Yıl Seçiniz:", choices = NULL),
    selectInput(inputId = "secilen_donem", label = "Çeyrek Seçiniz:", choices = NULL),
    radioButtons(inputId = "secilen_muhasebe", label = "Muhasebe Tipi:", choices = c("Normal", "Enflasyonlu"), selected = "Normal"),
    selectInput(inputId = "secilen_tur", label = "Şirket Türü:", choices = "Tümü"),
    hr(),
    
    conditionalPanel(
      condition = "input.ana_menu == 'kurum' || input.ana_menu == 'sektor'",
      h4("  Analiz Ayarları", style = "color: white; padding-left: 15px;"),
      
      conditionalPanel(
        condition = "input.ana_menu == 'kurum'",
        # YENİ: Arama çubuklu pickerInput (selectInput yerine)
        pickerInput(
          inputId = "secilen_kurum", 
          label = "Aracı Kurum Seçiniz:", 
          choices = NULL,
          options = list(
            `live-search` = TRUE,
            `live-search-placeholder` = "Kurum adını yazın...",
            size = 10
          )
        )
      ),
      
      selectInput(inputId = "grafik_metrik", label = "Analiz Edilecek Metrik:", 
                  choices = list(
                    "Bilanço ve Kârlılık" = c("Net Dönem Kârı" = "net_kar", "Özkaynak Büyüklüğü" = "iii_ozkaynaklar", "Toplam Varlıklar" = "toplam_varliklar"),
                    "Gelir Kalemleri" = c("Toplam Aracılık Geliri" = "toplam_net_aracilik_gelirleri", "Pay Piyasası Komisyonu" = "hisse_senedi_aracilik_gelirleri", "VİOP Komisyonu" = "turev_islemler_aracilik_gelirleri", "Yabancı Menkul Komisyonu" = "yabanci_menkul_kiymet_aracilik_gelirleri", "Kurumsal Finansman" = "kurumsal_finansman_gelirleri", "Faiz Gelirleri" = "faiz_gelirleri"),
                    "İşlem Hacimleri" = c("Pay Piyasası Hacmi" = "pay_islem_hacmi", "VİOP Hacmi" = "vadeli_islemler"),
                    "Operasyon ve Verimlilik" = c("Personel Sayısı" = "personel_sayisi", "Şube Ağı" = "sube_agi", "Personel Giderleri" = "personel_ucret_ve_giderleri"),
                    "Finansal Rasyolar" = c("Özsermaye Kârlılığı (ROE) %" = "roe", "Personel Başına Kâr" = "personel_basina_kar", "Kişi Başı Personel Gideri" = "kisi_basi_personel_gideri")
                  ), selected = "net_kar")
    )
  ),
  
  dashboardBody(
    tabItems(
      # --- SEKTÖR GÖRÜNÜMÜ SEKMESİ ---
      tabItem(tabName = "sektor",
              fluidRow(
                valueBoxOutput("sektor_kpi_kar", width = 3),
                valueBoxOutput("sektor_kpi_varlik", width = 3),
                valueBoxOutput("sektor_kpi_personel", width = 3),
                valueBoxOutput("sektor_kpi_kurum", width = 3)
              ),
              fluidRow(
                box(title = "Pazar Payı (İlk 5 ve Diğerleri)", status = "primary", solidHeader = TRUE, width = 5, plotlyOutput("pazar_payi_grafik", height = "350px")),
                box(title = "Tam Pazar Payı Dağılımı (%)", status = "info", solidHeader = TRUE, width = 7, DTOutput("pazar_payi_tablo", height = "350px"))
              ),
              fluidRow(
                box(title = "Tüm Kurumlar Liderlik Tablosu (Aşağı Kaydırın)", status = "success", solidHeader = TRUE, width = 6, 
                    div(style = "height: 400px; overflow-y: scroll; overflow-x: hidden;", 
                        plotlyOutput("liderler_grafik", height = "1200px"))),
                box(title = "Büyüme Şampiyonları (Geçen Yıla Göre % Artış - Top 10)", status = "warning", solidHeader = TRUE, width = 6, plotlyOutput("buyume_sampiyonlari_grafik", height = "400px"))
              )
      ),
      
      # --- KURUM ANALİZİ SEKMESİ ---
      tabItem(tabName = "kurum",
              fluidRow(
                column(width = 12,
                       fluidRow(
                         valueBoxOutput("kpi_varliklar", width = 3),
                         valueBoxOutput("kpi_kar", width = 3),
                         valueBoxOutput("kpi_gelir", width = 3),
                         valueBoxOutput("kpi_personel", width = 3)
                       ),
                       fluidRow(
                         box(width = 12, status = "primary",
                             h4(textOutput("siralama_metni"), style="margin-top:0px; font-weight:bold; text-align:center; color:#0073B7;")
                         )
                       ),
                       fluidRow(
                         box(title = "Kurum Performans Trendi", status = "info", solidHeader = TRUE, width = 6, plotlyOutput("trend_grafik", height = "320px")),
                         box(title = "Sektör Ortalaması (±1 Std. Sapma Kanalı)", status = "warning", solidHeader = TRUE, width = 6, plotlyOutput("karsilastirma_grafik", height = "320px"))
                       ),
                       fluidRow(
                         box(title = "Yıllık Büyüme İvmesi (Geçen Yılın Aynı Çeyreğine Göre)", status = "success", solidHeader = TRUE, width = 6, plotlyOutput("buyume_grafik", height = "350px")),
                         box(title = "Gelir Modeli Dağılımı (Milyon ₺)", status = "primary", solidHeader = TRUE, width = 6, plotlyOutput("gelir_dagilimi_grafik", height = "350px"))
                       ),
                       fluidRow(
                         box(title = "Yıllara Göre Çeyreklik Performans Karşılaştırması", status = "primary", solidHeader = TRUE, width = 12, plotlyOutput("ceyreklik_karsilastirma_grafik", height = "350px"))
                       ),
                       fluidRow(
                         box(title = "Kurum Sektör Profili (Yüzdelik Dilim)", status = "danger", solidHeader = TRUE, width = 4, plotlyOutput("radar_grafik", height = "400px")),
                         box(title = "Seçili Dönem Bilanço ve Gelir Tablosu Detayı", status = "info", solidHeader = TRUE, width = 8, DTOutput("veri_tablosu", height = "400px"))
                       )
                )
              )
      ),
      
      # --- MAKİNE ÖĞRENMESİ SEKMESİ ---
      tabItem(tabName = "ml", h2("Gelecek Dönem Kâr ve Hacim Tahminleri"))
    )
  )
)


# --- 2. SERVER (ARKA PLAN İŞLEMLERİ) ---
server <- function(input, output, session) {
  
  # --- A. VERİ GÜNCELLEME VE TEMİZLEME BORU HATTI ---
  veri_guncelleme_tetikleyicisi <- reactiveVal(0)
  
  observeEvent(input$yeni_excel, {
    req(input$yeni_excel)
    file.copy(from = input$yeni_excel$datapath, to = "veri.xlsx", overwrite = TRUE)
    veri_guncelleme_tetikleyicisi(veri_guncelleme_tetikleyicisi() + 1)
    showNotification("Veritabanı güncellendi: Yeni Excel başarıyla sisteme işlendi!", type = "message", duration = 5)
  })
  
  # Veri temizleme adımları Server'ın içine alındı (Dinamik çalışması için)
  ana_veri <- reactive({
    veri_guncelleme_tetikleyicisi()
    
    # 1. Dosyayı oku (Excel varsa excel'i, yoksa R'daki mevcut "veri" objesini kullanır)
    if (file.exists("veri.xlsx")) {
      ham_veri <- read_excel("veri.xlsx")
    } else if (exists("veri")) {
      ham_veri <- get("veri") 
    } else {
      return(data.frame()) # İkisi de yoksa çökmeyi önler
    }
    
    # 2. Temizlik ve formatlama (Orijinal kodundaki algoritmalar)
    ham_veri %>%
      clean_names() %>%
      remove_empty("rows") %>% 
      mutate(across(where(is.character), trimws)) %>% 
      mutate(across(where(is.numeric), ~replace_na(.x, 0))) %>% 
      mutate(across(where(is.character), ~replace_na(.x, "0"))) %>% 
      rename(net_kar = p6_donem_kari_zarari) %>%
      mutate(
        roe = ifelse(iii_ozkaynaklar != 0, (net_kar / iii_ozkaynaklar) * 100, 0),
        personel_basina_kar = ifelse(personel_sayisi > 0, net_kar / personel_sayisi, 0),
        kisi_basi_personel_gideri = ifelse(personel_sayisi > 0, personel_ucret_ve_giderleri / personel_sayisi, 0),
        zaman = paste(yil, donem, sep = "-")
      ) %>%
      arrange(yil, donem)
  })
  
  # --- B. FİLTRELERİ DİNAMİK OLARAK DOLDURMA ---
  observe({
    df <- ana_veri()
    req(nrow(df) > 0)
    
    updateSelectInput(session, "secilen_yil", choices = sort(unique(df$yil), decreasing = TRUE))
    updateSelectInput(session, "secilen_donem", choices = sort(unique(df$donem)))
    updateSelectInput(session, "secilen_tur", choices = c("Tümü", unique(df$sirket_turu)))
  })
  
  # --- C. MEVCUT VERİ BORULARI (Artık ana_veri üzerinden besleniyor) ---
  veri_global <- reactive({
    req(ana_veri())
    temp_veri <- ana_veri() %>% filter(muhasebe_tipi == input$secilen_muhasebe)
    if (input$secilen_tur != "Tümü") { temp_veri <- temp_veri %>% filter(sirket_turu == input$secilen_tur) }
    return(temp_veri)
  })
  
  # Kurum Picker Input'unu Güncelleme
  observe({
    req(veri_global())
    guncel_kurumlar <- sort(unique(veri_global()$unvani))
    mevcut_secim <- input$secilen_kurum
    yeni_secim <- if (!is.null(mevcut_secim) && mevcut_secim %in% guncel_kurumlar) { mevcut_secim } else { guncel_kurumlar[1] }
    
    updatePickerInput(session, "secilen_kurum", choices = guncel_kurumlar, selected = yeni_secim)
  })
  
  veri_kurum <- reactive({
    req(input$secilen_kurum, veri_global()) 
    veri_global() %>% filter(unvani == input$secilen_kurum)
  })
  
  veri_kpi <- reactive({
    veri_kurum() %>% filter(yil == input$secilen_yil, donem == input$secilen_donem)
  })
  
  sektor_verisi <- reactive({
    veri_global() %>% filter(yil == input$secilen_yil, donem == input$secilen_donem)
  })
  
  # YARDIMCI FORMATLAMA FONKSİYONU
  grafik_ayarlari <- function(metrik_kodu) {
    metrikler_adet <- c("personel_sayisi", "sube_agi")
    metrikler_yuzde <- c("roe")
    metrikler_kisi_basi <- c("personel_basina_kar", "kisi_basi_personel_gideri")
    
    isim_listesi <- c(
      "net_kar" = "Net Dönem Kârı", "iii_ozkaynaklar" = "Özkaynak Büyüklüğü", "toplam_varliklar" = "Toplam Varlıklar",
      "toplam_net_aracilik_gelirleri" = "Toplam Aracılık Geliri", "hisse_senedi_aracilik_gelirleri" = "Pay Piyasası Kom.",
      "turev_islemler_aracilik_gelirleri" = "VİOP Komisyonu", "yabanci_menkul_kiymet_aracilik_gelirleri" = "Yabancı Menkul Kom.",
      "kurumsal_finansman_gelirleri" = "Kurumsal Finansman", "faiz_gelirleri" = "Faiz Gelirleri",
      "pay_islem_hacmi" = "Pay Piyasası Hacmi", "vadeli_islemler" = "VİOP Hacmi",
      "personel_sayisi" = "Personel Sayısı", "sube_agi" = "Şube Ağı", "personel_ucret_ve_giderleri" = "Personel Giderleri",
      "roe" = "Özsermaye Kârlılığı (ROE)", "personel_basina_kar" = "Personel Başına Kâr", "kisi_basi_personel_gideri" = "Kişi Başı Personel Gideri"
    )
    
    metrik_adi <- isim_listesi[metrik_kodu]
    if (metrik_kodu %in% metrikler_adet) { list(adi = metrik_adi, bolen = 1, baslik = metrik_adi, ek = "", kusurat = 0)
    } else if (metrik_kodu %in% metrikler_yuzde) { list(adi = metrik_adi, bolen = 1, baslik = paste(metrik_adi, "(%)"), ek = " %", kusurat = 2)
    } else if (metrik_kodu %in% metrikler_kisi_basi) { list(adi = metrik_adi, bolen = 1, baslik = paste(metrik_adi, "(TL)"), ek = " ₺", kusurat = 0)
    } else { list(adi = metrik_adi, bolen = 1e6, baslik = paste(metrik_adi, "(Milyon TL)"), ek = " ₺", kusurat = 0) }
  }
  
  # ==========================================
  # 1. SEKTÖR GÖRÜNÜMÜ ÇIKTILARI
  # ==========================================
  
  output$sektor_kpi_kar <- renderValueBox({
    df <- sektor_verisi()
    toplam_kar <- sum(df$net_kar, na.rm = TRUE)
    gosterim <- ifelse(abs(toplam_kar) >= 1e9, paste0(format(round(toplam_kar / 1e9, 2), big.mark=".", decimal.mark=","), " Milyar ₺"), paste0(format(round(toplam_kar / 1e6, 1), big.mark=".", decimal.mark=","), " Milyon ₺"))
    valueBox(gosterim, "Sektör Toplam Kâr", icon = icon("coins"), color = "blue")
  })
  
  output$sektor_kpi_varlik <- renderValueBox({
    df <- sektor_verisi()
    toplam_varlik <- sum(df$toplam_varliklar, na.rm = TRUE)
    gosterim <- ifelse(abs(toplam_varlik) >= 1e9, paste0(format(round(toplam_varlik / 1e9, 2), big.mark=".", decimal.mark=","), " Milyar ₺"), paste0(format(round(toplam_varlik / 1e6, 1), big.mark=".", decimal.mark=","), " Milyon ₺"))
    valueBox(gosterim, "Sektör Toplam Varlık", icon = icon("money-bill-trend-up"), color = "green")
  })
  
  output$sektor_kpi_personel <- renderValueBox({
    df <- sektor_verisi()
    toplam_per <- sum(df$personel_sayisi, na.rm = TRUE)
    valueBox(format(toplam_per, big.mark="."), "Toplam İstihdam", icon = icon("users"), color = "aqua")
  })
  
  output$sektor_kpi_kurum <- renderValueBox({
    df <- sektor_verisi()
    valueBox(nrow(df), "Aktif Kurum Sayısı", icon = icon("building"), color = "purple")
  })
  
  output$pazar_payi_grafik <- renderPlotly({
    df <- sektor_verisi()
    if(nrow(df) == 0) return(NULL)
    
    metrik <- input$grafik_metrik
    gider_metrikleri <- c("personel_ucret_ve_giderleri", "kisi_basi_personel_gideri")
    
    df <- df %>%
      mutate(deger = if(metrik %in% gider_metrikleri) abs(.data[[metrik]]) else .data[[metrik]]) %>%
      filter(deger > 0) %>% arrange(desc(deger))
    
    toplam_pazar <- sum(df$deger, na.rm = TRUE)
    if(toplam_pazar == 0) return(NULL)
    
    ilk_5 <- df %>% slice(1:5) %>% select(unvani, deger)
    diger <- data.frame(unvani = "Diğer", deger = sum(df$deger[6:nrow(df)], na.rm = TRUE))
    pasta_veri <- bind_rows(ilk_5, diger) %>% mutate(yuzde = (deger / toplam_pazar) * 100)
    
    plot_ly(pasta_veri, labels = ~unvani, values = ~deger, type = 'pie', hole = 0.5,
            textposition = 'inside', 
            textinfo = 'percent', 
            hoverinfo = 'text', 
            text = ~paste0("<b>", unvani, "</b><br>Pazar Payı: %", format(round(yuzde, 1), decimal.mark = ","))) %>%
      layout(showlegend = TRUE, 
             legend = list(orientation = "v", x = 1.1, y = 0.5), 
             margin = list(t = 10, b = 10, l = 10, r = 130)) %>% 
      config(displayModeBar = FALSE)
  })
  
  output$pazar_payi_tablo <- renderDT({
    df <- sektor_verisi()
    if(nrow(df) == 0) return(NULL)
    
    metrik <- input$grafik_metrik
    gider_metrikleri <- c("personel_ucret_ve_giderleri", "kisi_basi_personel_gideri")
    
    df_tablo <- df %>%
      mutate(Ham_Deger = if(metrik %in% gider_metrikleri) abs(.data[[metrik]]) else .data[[metrik]]) %>%
      filter(Ham_Deger > 0) %>%
      mutate(
        Pazar_Payi = (Ham_Deger / sum(Ham_Deger, na.rm = TRUE)) * 100,
        Deger_Gosterim = format(round(Ham_Deger, 0), big.mark = ".", scientific = FALSE, trim = TRUE)
      ) %>%
      arrange(desc(Ham_Deger)) %>%
      select(`Aracı Kurum` = unvani, `Değer` = Deger_Gosterim, `Pazar Payı (%)` = Pazar_Payi) %>%
      mutate(`Pazar Payı (%)` = round(`Pazar Payı (%)`, 2))
    
    datatable(df_tablo, options = list(pageLength = 150, scrollY = "200px", dom = 'ft'), rownames = FALSE, selection = "none") %>%
      formatStyle('Pazar Payı (%)', background = styleColorBar(c(0, max(df_tablo$`Pazar Payı (%)`)), 'lightblue'), backgroundSize = '98% 88%', backgroundRepeat = 'no-repeat', backgroundPosition = 'center')
  })
  
  output$liderler_grafik <- renderPlotly({
    df <- sektor_verisi()
    if(nrow(df) == 0) return(NULL)
    
    metrik <- input$grafik_metrik
    ayar <- grafik_ayarlari(metrik)
    gider_metrikleri <- c("personel_ucret_ve_giderleri", "kisi_basi_personel_gideri")
    
    df <- df %>% mutate(deger = if(metrik %in% gider_metrikleri) abs(.data[[metrik]]) else .data[[metrik]]) %>% arrange(deger) 
    
    df$unvani <- factor(df$unvani, levels = df$unvani)
    df$y_degeri <- df$deger / ayar$bolen
    df$hover_metni <- paste0("Kurum: ", df$unvani, "<br>", ayar$adi, ": ", format(round(df$deger, ayar$kusurat), big.mark = ".", decimal.mark = ",", scientific = FALSE), ayar$ek)
    
    p <- ggplot(df, aes(x = unvani, y = y_degeri, text = hover_metni)) +
      geom_col(fill = "#0073B7", width = 0.7) +
      coord_flip() + theme_minimal() + labs(x = NULL, y = ayar$baslik) + theme(axis.text.y = element_text(size = 9))
    
    suppressWarnings(ggplotly(p, tooltip = "text") %>% config(displayModeBar = FALSE) %>% layout(margin = list(l = 150))) 
  })
  
  output$buyume_sampiyonlari_grafik <- renderPlotly({
    df_tumu <- veri_global() 
    if(nrow(df_tumu) == 0) return(NULL)
    
    metrik <- input$grafik_metrik
    gider_metrikleri <- c("personel_ucret_ve_giderleri", "kisi_basi_personel_gideri")
    
    df_tumu <- df_tumu %>% mutate(deger = if(metrik %in% gider_metrikleri) abs(.data[[metrik]]) else .data[[metrik]])
    
    df_hesap <- df_tumu %>%
      arrange(unvani, donem, yil) %>% group_by(unvani, donem) %>%
      mutate(
        gecen_yil_deger = lag(deger, order_by = yil),
        buyume_orani = ifelse(gecen_yil_deger > 0, ((deger - gecen_yil_deger) / gecen_yil_deger) * 100, NA)
      ) %>%
      ungroup() %>%
      filter(yil == input$secilen_yil, donem == input$secilen_donem) %>%
      drop_na(buyume_orani) %>% arrange(desc(buyume_orani)) %>% head(10) 
    
    if(nrow(df_hesap) == 0) return(NULL)
    
    df_hesap$unvani <- factor(df_hesap$unvani, levels = rev(df_hesap$unvani))
    df_hesap$hover_metni <- paste0("Kurum: ", df_hesap$unvani, "<br>Büyüme: %", format(round(df_hesap$buyume_orani, 1), big.mark = ".", decimal.mark = ","))
    
    p <- ggplot(df_hesap, aes(x = unvani, y = buyume_orani, text = hover_metni)) +
      geom_col(fill = "#2CA02C", width = 0.7) +
      coord_flip() + theme_minimal() + labs(x = NULL, y = "Geçen Yılın Aynı Çeyreğine Göre Büyüme (%)") + theme(axis.text.y = element_text(size = 9))
    
    suppressWarnings(ggplotly(p, tooltip = "text") %>% config(displayModeBar = FALSE) %>% layout(margin = list(l = 150)))
  })
  
  # ==========================================
  # 2. KURUM ANALİZİ ÇIKTILARI
  # ==========================================
  
  output$kpi_varliklar <- renderValueBox({
    kpi_data <- veri_kpi()
    if(nrow(kpi_data) == 0) return(valueBox("Veri Yok", "Toplam Varlıklar", icon = icon("ban"), color = "navy"))
    ham_deger <- kpi_data$toplam_varliklar
    gosterim <- ifelse(abs(ham_deger) >= 1e9, paste0(format(round(ham_deger / 1e9, 2), big.mark = ".", decimal.mark = ","), " Milyar ₺"), paste0(format(round(ham_deger / 1e6, 1), big.mark = ".", decimal.mark = ","), " Milyon ₺"))
    valueBox(gosterim, "Toplam Varlıklar", icon = icon("money-bill-trend-up"), color = "green")
  })
  
  output$kpi_kar <- renderValueBox({
    kpi_data <- veri_kpi()
    if(nrow(kpi_data) == 0) return(valueBox("Veri Yok", "Net Dönem Kârı", icon = icon("ban"), color = "navy"))
    ham_deger <- kpi_data$net_kar
    gosterim <- ifelse(abs(ham_deger) >= 1e9, paste0(format(round(ham_deger / 1e9, 2), big.mark = ".", decimal.mark = ","), " Milyar ₺"), paste0(format(round(ham_deger / 1e6, 1), big.mark = ".", decimal.mark = ","), " Milyon ₺"))
    valueBox(gosterim, "Net Dönem Kârı", icon = icon("coins"), color = ifelse(ham_deger > 0, "blue", "red"))
  })
  
  output$kpi_gelir <- renderValueBox({
    kpi_data <- veri_kpi()
    if(nrow(kpi_data) == 0) return(valueBox("Veri Yok", "Aracılık Geliri", icon = icon("ban"), color = "navy"))
    ham_deger <- kpi_data$toplam_net_aracilik_gelirleri
    gosterim <- ifelse(abs(ham_deger) >= 1e9, paste0(format(round(ham_deger / 1e9, 2), big.mark = ".", decimal.mark = ","), " Milyar ₺"), paste0(format(round(ham_deger / 1e6, 1), big.mark = ".", decimal.mark = ","), " Milyon ₺"))
    valueBox(gosterim, "Aracılık Geliri", icon = icon("chart-pie"), color = "purple")
  })
  
  output$kpi_personel <- renderValueBox({
    kpi_data <- veri_kpi()
    if(nrow(kpi_data) == 0) return(valueBox("Veri Yok", "Personel Sayısı", icon = icon("ban"), color = "navy"))
    valueBox(kpi_data$personel_sayisi, "Personel Sayısı", icon = icon("users"), color = "aqua")
  })
  
  output$siralama_metni <- renderText({
    req(input$secilen_kurum)
    c_data <- veri_global() %>% filter(yil == input$secilen_yil, donem == input$secilen_donem)
    if(nrow(c_data) == 0) return("Seçili dönem için sıralama verisi bulunmamaktadır.")
    
    ayar <- grafik_ayarlari(input$grafik_metrik)
    gider_metrikleri <- c("personel_ucret_ve_giderleri", "kisi_basi_personel_gideri")
    
    if (input$grafik_metrik %in% gider_metrikleri) {
      c_data <- c_data %>% arrange(desc(abs(.data[[input$grafik_metrik]]))) %>% mutate(sira = row_number())
      baglam_metni <- "en yüksek harcamayı/gideri yapan"
    } else {
      c_data <- c_data %>% arrange(desc(.data[[input$grafik_metrik]])) %>% mutate(sira = row_number())
      baglam_metni <- "en yüksek değere sahip"
    }
    
    kurum_sira <- c_data %>% filter(unvani == input$secilen_kurum) %>% pull(sira)
    toplam_kurum <- nrow(c_data)
    if(length(kurum_sira) == 0) return("Kurum bu dönemde listede yok.")
    
    paste0("📊 Sektör Sıralaması: ", input$secilen_kurum, ", ", input$secilen_yil, "-", input$secilen_donem, " döneminde '", ayar$adi, "' kaleminde ", baglam_metni, " ", toplam_kurum, " kurum arasında ", kurum_sira, ". sıradadır.")
  })
  
  output$trend_grafik <- renderPlotly({
    grafik_verisi <- veri_kurum()
    if(nrow(grafik_verisi) == 0) return(NULL) 
    ayar <- grafik_ayarlari(input$grafik_metrik)
    grafik_verisi$y_degeri <- grafik_verisi[[input$grafik_metrik]] / ayar$bolen
    grafik_verisi$hover_metni <- paste0("Dönem: ", grafik_verisi$zaman, "<br>", ayar$adi, ": ", format(round(grafik_verisi[[input$grafik_metrik]], ayar$kusurat), big.mark = ".", decimal.mark = ",", scientific = FALSE), ayar$ek)
    
    p <- ggplot(grafik_verisi, aes(x = zaman, y = y_degeri, group = 1, text = hover_metni)) +
      geom_line(color = "#0073B7", linewidth = 1) + geom_point(color = "#0073B7", size = 3) +
      theme_minimal() + labs(x = NULL, y = ayar$baslik) + theme(axis.text.x = element_text(angle = 45, hjust = 1))
    suppressWarnings(ggplotly(p, tooltip = "text") %>% config(displayModeBar = FALSE) %>% layout(margin = list(b = 40)))
  })
  
  output$karsilastirma_grafik <- renderPlotly({
    grafik_verisi <- veri_kurum()
    if(nrow(grafik_verisi) == 0) return(NULL) 
    ayar <- grafik_ayarlari(input$grafik_metrik)
    
    sektor_ortalamasi <- veri_global() %>% group_by(zaman) %>%
      summarise(ortalama_deger = mean(.data[[input$grafik_metrik]], na.rm = TRUE), sd_deger = sd(.data[[input$grafik_metrik]], na.rm = TRUE), .groups = "drop")
    
    birlesik_veri <- grafik_verisi %>% left_join(sektor_ortalamasi, by = "zaman")
    birlesik_veri$y_degeri <- birlesik_veri[[input$grafik_metrik]] / ayar$bolen
    birlesik_veri$y_ortalama <- birlesik_veri$ortalama_deger / ayar$bolen
    birlesik_veri$y_ust_bant <- (birlesik_veri$ortalama_deger + birlesik_veri$sd_deger) / ayar$bolen
    birlesik_veri$y_alt_bant <- (birlesik_veri$ortalama_deger - birlesik_veri$sd_deger) / ayar$bolen
    
    birlesik_veri$hover_metni <- paste0("Dönem: ", birlesik_veri$zaman, "<br>", "Kurum: ", input$secilen_kurum, "<br>", ayar$adi, ": ", format(round(birlesik_veri[[input$grafik_metrik]], ayar$kusurat), big.mark = ".", decimal.mark = ",", scientific = FALSE), ayar$ek)
    birlesik_veri$hover_ortalama <- paste0("Dönem: ", birlesik_veri$zaman, "<br>", "Sektör Ortalaması<br>", ayar$adi, ": ", format(round(birlesik_veri$ortalama_deger, ayar$kusurat), big.mark = ".", decimal.mark = ",", scientific = FALSE), ayar$ek)
    
    p <- ggplot(birlesik_veri, aes(x = zaman)) +
      geom_ribbon(aes(ymin = y_alt_bant, ymax = y_ust_bant, fill = "±1 Std. Sapma"), group = 1, alpha = 0.15) +
      geom_line(aes(y = y_ortalama, color = "Sektör Ortalaması", text = hover_ortalama), group = 1, linetype = "dashed", linewidth = 1) +
      geom_point(aes(y = y_ortalama, color = "Sektör Ortalaması", text = hover_ortalama), size = 2, alpha = 0.7) +
      geom_line(aes(y = y_degeri, color = "Seçili Kurum", text = hover_metni), group = 1, linewidth = 1.2) +
      geom_point(aes(y = y_degeri, color = "Seçili Kurum", text = hover_metni), size = 3) +
      scale_color_manual(name = NULL, values = c("Seçili Kurum" = "#E08E79", "Sektör Ortalaması" = "#808080")) +
      scale_fill_manual(name = NULL, values = c("±1 Std. Sapma" = "#2c3e50")) +
      theme_minimal() + labs(x = NULL, y = ayar$baslik) +
      theme(axis.text.x = element_text(angle = 45, hjust = 1), legend.position = "bottom", legend.title = element_blank())
    
    suppressWarnings(ggplotly(p, tooltip = "text") %>% config(displayModeBar = FALSE) %>% layout(legend = list(orientation = "h", x = 0.5, xanchor = "center", y = -0.2), margin = list(b = 50)))
  })
  
  output$buyume_grafik <- renderPlotly({
    grafik_verisi <- veri_kurum()
    if(nrow(grafik_verisi) < 2) return(NULL) 
    
    ayar <- grafik_ayarlari(input$grafik_metrik)
    
    grafik_verisi <- grafik_verisi %>%
      arrange(donem, yil) %>% group_by(donem) %>%
      mutate(
        gecen_yil_degeri = lag(.data[[input$grafik_metrik]]),
        degisim = .data[[input$grafik_metrik]] - gecen_yil_degeri,
        yon = ifelse(degisim > 0, "Artış", "Azalış")
      ) %>% ungroup() %>% arrange(yil, donem) %>% drop_na(degisim)
    
    grafik_verisi$y_degeri <- grafik_verisi$degisim / ayar$bolen
    grafik_verisi$hover_metni <- paste0("Dönem: ", grafik_verisi$zaman, "<br>", "Geçen Yıla Göre Değişim: ", format(round(grafik_verisi$degisim, ayar$kusurat), big.mark = ".", decimal.mark = ",", scientific = FALSE), ayar$ek)
    
    p <- ggplot(grafik_verisi, aes(x = zaman, y = y_degeri, fill = yon, text = hover_metni)) +
      geom_col(alpha = 0.8) +
      scale_fill_manual(values = c("Artış" = "#2CA02C", "Azalış" = "#D62728")) +
      theme_minimal() + labs(x = NULL, y = paste("Değişim", "(", ayar$baslik, ")")) + theme(axis.text.x = element_text(angle = 45, hjust = 1), legend.position = "none")
    
    suppressWarnings(ggplotly(p, tooltip = "text") %>% config(displayModeBar = FALSE) %>% layout(margin = list(b = 40)))
  })
  
  output$gelir_dagilimi_grafik <- renderPlotly({
    g_veri <- veri_kurum()
    if(nrow(g_veri) == 0) return(NULL)
    
    g_veri_uzun <- g_veri %>%
      select(zaman, `Pay Komisyonu` = hisse_senedi_aracilik_gelirleri, `VİOP Komisyonu` = turev_islemler_aracilik_gelirleri, `Faiz Geliri` = faiz_gelirleri, `Kurumsal Finansman` = kurumsal_finansman_gelirleri, `Yabancı Menkul` = yabanci_menkul_kiymet_aracilik_gelirleri) %>%
      pivot_longer(cols = -zaman, names_to = "gelir_turu", values_to = "deger") %>% filter(deger > 0)
    
    g_veri_uzun$y_degeri <- g_veri_uzun$deger / 1e6
    g_veri_uzun$hover_metni <- paste0("Dönem: ", g_veri_uzun$zaman, "<br>", g_veri_uzun$gelir_turu, ": ", format(round(g_veri_uzun$deger, 0), big.mark = ".", decimal.mark = ","), " ₺")
    
    p <- ggplot(g_veri_uzun, aes(x = zaman, y = y_degeri, fill = gelir_turu, text = hover_metni)) +
      geom_col(position = "stack") + scale_fill_brewer(palette = "Set2") +
      theme_minimal() + labs(x = NULL, y = "Milyon TL", fill = "") + theme(axis.text.x = element_text(angle = 45, hjust = 1), legend.position = "bottom")
    
    suppressWarnings(ggplotly(p, tooltip = "text") %>% config(displayModeBar = FALSE) %>% layout(legend = list(orientation = "h", x = 0.5, xanchor = "center", y = -0.2), margin = list(b = 50)))
  })
  
  output$ceyreklik_karsilastirma_grafik <- renderPlotly({
    df <- veri_kurum()
    if(nrow(df) == 0) return(NULL)
    
    ayar <- grafik_ayarlari(input$grafik_metrik)
    df$y_degeri <- df[[input$grafik_metrik]] / ayar$bolen
    df$hover_metni <- paste0("Yıl: ", df$yil, " - Çeyrek: ", df$donem, "<br>", ayar$adi, ": ", format(round(df[[input$grafik_metrik]], ayar$kusurat), big.mark = ".", decimal.mark = ",", scientific = FALSE), ayar$ek)
    
    p <- ggplot(df, aes(x = donem, y = y_degeri, fill = as.factor(yil), text = hover_metni)) +
      geom_col(position = position_dodge(width = 0.8), width = 0.7, alpha = 0.9) +
      scale_fill_brewer(palette = "Blues") + 
      theme_minimal() + 
      labs(x = NULL, y = ayar$baslik, fill = "Yıl:") + 
      theme(axis.text.x = element_text(size = 11, face = "bold"))
    
    suppressWarnings(ggplotly(p, tooltip = "text") %>% 
                       config(displayModeBar = FALSE) %>% 
                       layout(legend = list(orientation = "h", x = 0.5, xanchor = "center", y = -0.25), 
                              margin = list(b = 60)))
  })
  
  output$radar_grafik <- renderPlotly({
    req(input$secilen_kurum)
    c_data <- veri_global() %>% filter(yil == input$secilen_yil, donem == input$secilen_donem)
    if(nrow(c_data) < 2) return(NULL)
    
    calc_perc <- function(x) { round(rank(x, ties.method = "average") / length(x) * 100, 1) }
    
    c_data <- c_data %>% mutate(perc_kar = calc_perc(net_kar), perc_ozkaynak = calc_perc(iii_ozkaynaklar), perc_varlik = calc_perc(toplam_varliklar), perc_gelir = calc_perc(toplam_net_aracilik_gelirleri))
    kurum_data <- c_data %>% filter(unvani == input$secilen_kurum)
    if(nrow(kurum_data) == 0) return(NULL)
    
    fig <- plot_ly(
      type = 'scatterpolar',
      r = c(kurum_data$perc_kar, kurum_data$perc_ozkaynak, kurum_data$perc_varlik, kurum_data$perc_gelir, kurum_data$perc_kar),
      theta = c('Net Kâr', 'Özkaynak', 'Toplam Varlık', 'Aracılık Geliri', 'Net Kâr'),
      fill = 'toself', fillcolor = 'rgba(0, 115, 183, 0.4)', line = list(color = '#0073B7', width = 2),
      marker = list(color = '#0073B7', size = 8), hoverinfo = "text",
      text = paste0("%", c(kurum_data$perc_kar, kurum_data$perc_ozkaynak, kurum_data$perc_varlik, kurum_data$perc_gelir, kurum_data$perc_kar), " Dilim")
    ) 
    
    fig %>% layout(polar = list(radialaxis = list(visible = TRUE, range = c(0, 100))), showlegend = FALSE, margin = list(t = 20, b = 20, l = 90, r = 90)) %>% config(displayModeBar = FALSE)
  })
  
  output$veri_tablosu <- renderDT({
    req(input$secilen_kurum)
    df <- veri_kpi()
    if(nrow(df) == 0) return(NULL)
    
    df_uzun <- df %>%
      select(-zaman, -unvani, -sirket_turu, -muhasebe_tipi, -yil, -donem) %>%
      pivot_longer(cols = everything(), names_to = "Finansal Kalem", values_to = "Değer") %>%
      mutate(
        `Finansal Kalem` = toupper(gsub("_", " ", `Finansal Kalem`)), 
        Değer = suppressWarnings(as.numeric(Değer)),
        `Formatlı Değer` = format(round(Değer, 0), big.mark = ".", scientific = FALSE, trim = TRUE)
      ) %>%
      select(`Finansal Kalem`, `Değer` = `Formatlı Değer`)
    
    datatable(df_uzun, options = list(pageLength = 150, scrollY = "300px", dom = 'ft', language = list(search = "Kalem Ara: ", zeroRecords = "Kayıt bulunamadı.")), rownames = FALSE, selection = "none")
  })
}

# Uygulamayı Çalıştır
shinyApp(ui, server)

library("rsconnect")
install.packages("remotes")
