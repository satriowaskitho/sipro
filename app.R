library(bs4Dash)
library(writexl)
library(dplyr)
library(DT)
library(zip)
library(rlang)
library(openxlsx)
library(shinyjs)
library(shinyalert)
library(haven)
library(googlesheets4)



# Autentikasi pakai service account format .json
# gs4_auth(path = "/srv/shiny-server/sipro/gs4-sa.json")
gs4_auth(path = "credentials/gs4-sa.json")

options(scipen = 999)
R.version.string

ui <- bs4DashPage(
  
  title = "SIPRO", 
  bs4DashNavbar(title = "SIPRO", rightUi = NULL),
  
  
  bs4DashSidebar(
    useShinyjs(),
    tags$style(HTML("
    
     .navbar-nav.ml-auto {
      display: none !important;
    }
    
      .main-sidebar {
        position: fixed;
        width: 250px;
        overflow-y: auto;
        }
    
        .blurred {
        filter: blur(5px);
        pointer-events: none;
        user-select: none;
      }"
    )),
    
    sidebarMenu(
      menuItem(HTML("&nbsp;&nbsp;Beranda"), tabName = "dashboard", icon = icon("home")),
      menuItem(HTML("&nbsp;&nbsp;Preprocessing Data"), tabName = "preprocessing_data", icon = icon("cogs")),
      menuItem(HTML("&nbsp;&nbsp;Olah Data"), icon = icon("chart-column"),startExpanded = TRUE,
               menuSubItem(HTML("&nbsp;&nbsp;Ekspor"), tabName = "ekspor",icon = icon("plane-departure")),
               menuSubItem(HTML("&nbsp;&nbsp;Impor"), tabName = "impor", icon = icon("plane-arrival")),
               menuSubItem(HTML("&nbsp;&nbsp;Neraca Perdagangan"), tabName = "neraca", icon = icon("balance-scale"))),
      menuItem(HTML("&nbsp;&nbsp;Kalender Kegiatan"), tabName = "calendar_events", icon = icon("calendar-alt")),
      menuItem(HTML("&nbsp;&nbsp;FAQ"),tabName = "faq", icon = icon("comments"))
    ),
    tags$div(
      style = "width: 100%; text-align: center; padding: 20px 0;",
      tags$img(src = "logo_sipro.png", style = "max-width: 80%; height: auto;margin-left:-10px;")
    )
    
  ),
  
  bs4DashBody(
    useShinyjs(),
    tags$head(
      tags$link(rel = "shortcut icon", href = "logo_sipro.png"),
      tags$style(HTML("
      
      
.modal.fade .modal-dialog {
  opacity: 0;
  transform: translateY(-10px); /* lebih ringan dari -30px */
  transition: opacity 0.2s ease-out, transform 0.2s ease-out; /* lebih cepat */
}

.modal.fade.show .modal-dialog {
  opacity: 1;
  transform: translateY(0);
}

.modal-backdrop.fade {
  opacity: 0;
  transition: opacity 0.2s ease-out;
}

.modal-backdrop.show {
  opacity: 0.4; /* bisa disesuaikan */
}

      
          /* Pastikan semua container tidak melebihi lebar layar */
    .container, .content-wrapper, .right-side, .main-sidebar, .content {
      max-width: 100vw !important;
      overflow-x: hidden !important;
    }

    /* Untuk menghindari horizontal scroll di layar kecil */
    body {
      overflow-x: hidden !important;
    }

    

    /* Optional: agar modal content tidak berjarak aneh */
    .modal-content {
      padding: 10px;
    }
    
      
      body, label, input, button, select, .dataTables_wrapper {
      font-size: 14px !important; /* Atur sesuai kebutuhan */
    }

    table.dataTable td, table.dataTable th {
      font-size: 14px !important;
    }
      
        .btn-primary {
      background-color: #007bff !important;
      border-color: #007bff !important;
      color: white !important;
    }
    .btn-primary:hover {
      background-color: #0056b3 !important;
      border-color: #0056b3 !important;
    }

      
    /* Pastikan header tetap di atas */
    .main-header {
      position: sticky;
      top: 0;
      z-index: 1030;
    }

    /* Pastikan isi body tidak naik ke bawah header */
    .content-wrapper {
  margin-left: 250px;
}

    /* Perbaiki footer agar tidak mengganggu layout responsif */
     .main-footer {
          background-color: white;
          height: 50px;
          border-top: 1px solid #ddd;
          position: fixed;
          bottom: 0;
          left: 0;
          right: 0;
          padding: 20px;
          box-shadow: 0 -2px 5px rgba(0,0,0,0.1);
          margin-left: 250px;
        }

    /* Blur effect untuk konten terkunci */
    .blurred {
      filter: blur(5px);
      pointer-events: none;
      user-select: none;
    }

    /* Responsive fix agar layout tetap adaptif */
    @media (max-width: 768px) {
      .main-footer {
        margin-left: 0 !important;
      }
    }
  "))
    )
    ,
    div(id = "app_content",
        tabItems(
          tabItem("dashboard",
                  div(
                    id = "dashboard-content", 
                    div(
                      style = "width: 100%; background-color: white; padding: 8px 12px; border-radius: 0 0 8px 8px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); margin-bottom: 15px;",
                      tags$h4(
                        icon("home"), 
                        HTML("&nbsp;Beranda"),
                        style = "margin: 0;"
                      )
                    ),

                    fluidRow(
                      box(
                        title = tagList(icon("info"), HTML("&nbsp;&nbsp;Info Aplikasi")),
                        status = "primary",
                        width = 12,
                        solidHeader = TRUE,
                        collapsible = FALSE,
                        header = NULL,
                        style = "background-color: #f8f9fa; border: none; box-shadow: 0 2px 8px rgba(0, 0, 0, 0.05); padding: 24px; border-radius: 12px; margin-top: -10px;",
                        
                        # Judul
                        tags$h2(
                          HTML("Selamat Datang di <b style='color:#007BFF'>SIPRO</b>"),
                          style = "font-weight: bold; margin-top: 0; margin-bottom: 24px;"
                        ),
                        
                        # Konten utama
                        # Bagian gambar dan teks
                        fluidRow(
                          # Kolom gambar (kiri)
                          column(
                            width = 4,
                            tags$div(
                              style = "position: relative; height: 100%; min-height: 100px;",  # tingginya otomatis dari kolom kanan
                              tags$img(
                                src = "beranda.png",
                                style = "
          width: 100%; 
          max-width: 400px; 
          position: relative; 
          top: 50%; 
          transform: translateY(-50%);
          display: block; 
          margin: 0 auto;"
                              )
                            )
                          ),
                          
                          # Kolom teks dan tombol (kanan)
                          column(
                            width = 8,
                            tags$div(
                              style = "padding: 10px 20px; text-align: justify;",
                              
                              tags$p(HTML(
                                "<b>SIPRO</b> (Sistem Pengolahan Rutin Ekspor Impor) adalah platform digital yang dirancang untuk memudahkan proses pengolahan data ekspor dan impor bulanan di <b>BPS Kabupaten Karimun</b> secara terstruktur, cepat, dan tepat."
                              ), style = "margin-bottom: 12px;"),
                              
                              tags$p("Aplikasi ini dilengkapi dengan berbagai fitur utama berikut:", style = "margin-bottom: 4px;"),
                              
                              tags$ol(
                                style = "margin-top: 0px;",
                                tags$li(HTML("<b>Beranda:</b> berisi panduan penggunaan aplikasi, tautan data historis, daftar tabel output, link publikasi, dan referensi kode HS 8 digit.")),
                                tags$li(HTML("<b>Preprocessing data:</b> mencakup penggabungan (merged data) dan pembersihan data (cleaning data).")),
                                tags$li(HTML("<b>Olah data:</b> meliputi pengolahan data ekspor, impor, dan neraca perdagangan.")),
                                tags$li(HTML("<b>Kalender Kegiatan:</b> untuk memantau jadwal kegiatan terkait pengolahan dan publikasi data.")),
                                tags$li(HTML("<b>FAQ:</b> menampilkan daftar pertanyaan yang sering muncul serta menyediakan ruang untuk memberikan saran terkait pengembangan SIPRO ke depannya."))
                              ),
                              
                              # Tombol-tombol yang mengecil otomatis agar tetap sejajar
                              tags$div(
                                style = "
        display: flex;
        flex-wrap: nowrap;
        gap: 10px;
        margin-top: 20px;
        justify-content: space-between;
      ",
                                
                                tags$div(
                                  style = "flex: 1 1 auto;",
                                  actionButton(
                                    "btn_download_handbook",
                                    label = tagList(icon("download"), HTML("&nbsp;<small>Download Handbook</small>")),
                                    onclick = "window.open('Handbook_SIPRO.pdf', '_blank')",
                                    class = "btn btn-primary w-100",
                                    style = "font-size: 0.75rem; padding: 6px 4px;"
                                  )
                                ),
                                tags$div(
                                  style = "flex: 1 1 auto;",
                                  actionButton(
                                    "btn_video_tutorial",
                                    label = tagList(icon("youtube"), HTML("&nbsp;<small>Video Tutorial</small>")),
                                    onclick = "window.open('https://youtu.be/qcex52FtuCc', '_blank')",
                                    class = "btn btn-primary w-100",
                                    style = "font-size: 0.75rem; padding: 6px 4px;"
                                  )
                                ),
                                tags$div(
                                  style = "flex: 1 1 auto;",
                                  actionButton(
                                    "akses_data_historis",
                                    label = tagList(icon("database"), HTML("&nbsp;<small>Data Historis</small>")),
                                    class = "btn btn-primary w-100",
                                    style = "font-size: 0.75rem; padding: 6px 4px;"
                                  )
                                )
                              )
                            )
                          )
                          
                          
                        )
                        
                        
                      )
                    )
                    
                    
                    
                    ,
                    
                    fluidRow(
                      box(
                        title = tagList(icon("table"), HTML("&nbsp;&nbsp;Daftar Output")),
                        status = "primary",
                        solidHeader = TRUE,
                        collapsible = FALSE,
                        width = 12,
                        
                        tabBox(
                          id = "tab_output_view",
                          width = 12,  # sesuaikan dengan lebar container jika perlu
                          tabPanel(
                            title = "Tabel Output",
                            DTOutput("tabelOutput")
                          ),
                          tabPanel(
                            title = "Publikasi",
                            fluidRow(
                              column(
                                width = 4,
                                align = "center",
                                tags$img(
                                  src = "bahan_paparan.png",  # pastikan file ada di www/
                                  width = "100%",
                                  alt = "Ilustrasi Bahan Paparan"
                                )
                              ),
                              column(
                                width = 8,
                                tags$div(
                                  style = "padding: 10px;",
                                  tags$h4("Publikasi Ekspor dan Impor"),
                                  tags$p("Publikasi berupa Berita Resmi Statistik (BRS) serta bahan tayang ekspor dan impor bulanan secara periodik dapat diakses pada link berikut."),
                                  useShinyjs(),
                                  actionButton(
                                    "akses_paparan_button",
                                    label = HTML('<i class="fas fa-external-link-alt"></i> Lihat Publikasi'),
                                    class = "btn btn-primary"
                                  )
                                )
                              )
                              
                              
                            )
                          ),
                          tabPanel(
                            title = "Kode HS 8 Digit",
                            fluidPage(
                              fluidRow(
                                column(
                                  width = 4,
                                  align = "center",
                                  tags$img(
                                    src = "hscode.png",  # Pastikan file ini ada di folder www/
                                    width = "100%",
                                    alt = "Ilustrasi HS Code"
                                  )
                                ),
                                column(
                                  width = 8,
                                  tags$div(
                                    style = "padding: 10px;",
                                    tags$h4("Kode HS 8 Digit"),
                                    tags$p(
                                      "Daftar kode HS 8 digit yang digunakan pada data ekspor dan impor ini 
            mengacu pada HS Code Master yang berlaku sejak tahun 2022 hingga sekarang. 
            Terdapat sebanyak 11.554 kode yang digunakan untuk klasifikasi produk."
                                    ),
                                    tags$a(
                                      href = "HSCode Master BPS.pdf",
                                      target = "_blank",
                                      class = "btn btn-primary",
                                      HTML("<i class='fas fa-external-link-alt'></i>&nbsp;Lihat HS Code Master")
                                    )
                                  )
                                )
                              )
                            )
                          )
                          
                          
                          
                        )
                        
                      )
                    ))
                  
          ),
          
          
          tabItem("preprocessing_data",
                  div(
                    id = "dashboard-content",
                    div(
                      style = "width: 100%; background-color: white; padding: 8px 12px; border-radius: 0 0 8px 8px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); margin-bottom: 15px;",
                      tags$h4(
                        icon("cogs"), 
                        HTML("&nbsp;Preprocessing Data"),
                        style = "margin: 0;"
                      )
                    ),
                    
                    fluidRow(
                      column(
                        width = 12,
                        box(
                          title = tagList(icon("upload"), HTML("&nbsp;&nbsp;Input Data")),
                          status = "primary",
                          solidHeader = TRUE,
                          collapsible= FALSE,
                          width = NULL,  # supaya full lebar column
                          
                          numericInput("jumlah_file", "Banyak file yang ingin digabung", value = 2, min = 2),
                          tags$script(HTML("
  $(document).ready(function() {
    $('#jumlah_file').on('input change', function() {
      if (parseInt(this.value) < 2 || isNaN(this.value)) {
        this.value = 2;
      }
    });
  });
")),
                          uiOutput("file_inputs_ui"),
                          
                          tags$div(
                            style = "width: 100%;",
                            actionButton(
                              "clean_button", 
                              label = tagList(
                                icon("cogs", style = "color: white;"),
                                tags$span("Preprocessing Data", style = "color: white;")
                              ), 
                              class = "btn btn-primary", 
                              style = "width: 100%;"
                            )
                          )
                        ),
                        
                        br(),
                        
                        box(
                          title = tagList(icon("table"), HTML("&nbsp;&nbsp;Data Bersih")),
                          status = "primary",
                          solidHeader = TRUE,
                          collapsible= FALSE,
                          width = NULL,  # full lebar column
                          div(
                            style = "display: flex; flex-wrap: wrap; gap: 10px; justify-content: flex-start;",
                            downloadButton("downloadDataClean", "Download Data Bersih", style = "flex: 1 1 200px;")
                          ),
                          tags$br(), tags$br(),
                          DTOutput("tabel_output_clean_3")
                        )
                      )
                    )
                    
                  )),
          
          
          tabItem("ekspor",
                  div(
                    id = "dashboard-content",
                    div(
                      style = "width: 100%; background-color: white; padding: 8px 12px; border-radius: 0 0 8px 8px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); margin-bottom: 15px;",
                      tags$h4(
                        icon("plane-departure"), 
                        HTML("&nbsp;Olah Data Ekspor"),
                        style = "margin: 0;"
                      )
                    ),
                    
                    fluidRow(
                      box(
                        title = tagList(icon("upload"), HTML("&nbsp;&nbsp;Input Data Ekspor")),
                        status = "primary",
                        solidHeader = TRUE,
                        collapsible= FALSE,
                        width = 3,
                        fileInput("file_input", "Pilih Data Ekspor (.sav)", accept = ".sav"),
                        selectInput("tahun", "Pilih Tahun", choices = NULL),
                        selectInput("bulan", "Pilih Bulan", choices = NULL),
                        tags$div(
                          style = "width: 100%;",
                          actionButton(
                            "analisis_button", 
                            label = tagList(
                              icon("magnifying-glass-chart", style = "color: white;"),
                              tags$span("Analisis", style = "color: white;")
                            ), 
                            class = "btn btn-primary", 
                            style = "width: 100%;"
                          )
                        )
                        
                      ),
                      
                      box(
                        title = tagList(icon("table"), HTML("&nbsp;&nbsp;Tabel Output Ekspor")),
                        status = "primary",
                        solidHeader = TRUE,
                        collapsible= FALSE,
                        width = 9,
                        
                        tabBox(
                          width = 12,
                          
                          # Tab 1: Pilih Tabel + Tombol Download
                          tabPanel(
                            title = "Tabel Output",
                            selectInput(
                              inputId = "pilih_tabel",
                              label = NULL,
                              choices = list(
                                "Ringkasan Ekspor" = c(
                                  "1. Nilai Ekspor Menurut Sektor",
                                  "2. Nilai Ekspor Menurut Pelabuhan",
                                  "3. Perkembangan Nilai Ekspor",
                                  "4. Perkembangan Nilai Ekspor (c-t-c)",
                                  "5. Volume Ekspor Menurut Pelabuhan"
                                ),
                                "Ekspor Menurut Negara Tujuan" = c(
                                  "6. Nilai Ekspor Menurut Negara Tujuan",
                                  "7. Nilai Ekspor Nonmigas Menurut Negara Tujuan",
                                  "8. Nilai Ekspor Migas Menurut Negara Tujuan",
                                  "9. Nilai Ekspor Negara Tujuan Utama HS2 Digit",
                                  "10. Perkembangan Nilai Ekspor Negara Tujuan Utama",
                                  "11. Nilai Ekspor Kumulatif Menurut Negara Tujuan",
                                  "12. Perkembangan Nilai Ekspor Negara Tujuan Utama (c-t-c)"
                                ),
                                "Ekspor Menurut HS2 Digit" = c(
                                  "13. Nilai Ekspor Nonmigas Menurut Golongan Barang HS2 Digit",
                                  "14. Peningkatan/Penurunan Nilai Ekspor Nonmigas HS2 Digit (m-t-m)",
                                  "15. Share Nilai Ekspor Nonmigas HS2 Digit"
                                )
                              ),
                              selected = "1. Nilai Ekspor Menurut Sektor"
                            )
                            ,
                            div(
                              style = "display: flex; flex-wrap: wrap; gap: 10px; justify-content: flex-start;",
                              downloadButton("downloadData", "Download Tabel", style = "flex: 1 1 200px;"),
                              downloadButton("downloadAll", "Download Semua Tabel", style = "flex: 1 1 200px;")
                            ),
                            tags$br(),
                            DTOutput("tabel_output_ekspor_3")
                          )
                          ,
                          
                          # Tab 2: Tabel Output
                          tabPanel(
                            title = "Data",
                            DTOutput("tabel_data_ekspor")
                          )
                        )
                      )
                      
                    )
                    
                  )),
          
          tabItem("impor",
                  div(
                    id = "dashboard-content",
                    div(
                      style = "width: 100%; background-color: white; padding: 8px 12px; border-radius: 0 0 8px 8px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); margin-bottom: 15px;",
                      tags$h4(
                        icon("plane-arrival"), 
                        HTML("&nbsp;Olah Data Impor"),
                        style = "margin: 0;"
                      )
                    ),
                    
                    fluidRow(
                      box(
                        title = tagList(icon("upload"), HTML("&nbsp;&nbsp;Input Data Impor")),
                        status = "primary",
                        solidHeader = TRUE,
                        collapsible= FALSE,
                        width = 3,
                        fileInput("file_input_impor", "Pilih Data Impor (.sav)", accept = ".sav"),
                        selectInput("tahun_impor", "Pilih Tahun", choices = NULL),
                        selectInput("bulan_impor", "Pilih Bulan", choices = NULL),
                        tags$div(
                          style = "width: 100%;",
                          actionButton(
                            "analisis_button_impor", 
                            label = tagList(
                              icon("magnifying-glass-chart", style = "color: white;"),
                              tags$span("Analisis", style = "color: white;")
                            ), 
                            class = "btn btn-primary", 
                            style = "width: 100%;"
                          )
                        )
                        
                      ),
                      
                      box(
                        title = tagList(icon("table"), HTML("&nbsp;&nbsp;Tabel Output Impor")),
                        status = "primary",
                        solidHeader = TRUE,
                        collapsible= FALSE,
                        width = 9,
                        
                        tabBox(
                          width = 12,
                          
                          # Tab 1: Pilih Tabel + Tombol Download
                          tabPanel(
                            title = "Tabel Output",
                            selectInput(
                              inputId = "pilih_tabel_impor",
                              label = NULL,
                              choices = list(
                                "Ringkasan Impor" = c(
                                  "1. Nilai Impor Menurut Sektor",
                                  "2. Nilai Impor Menurut Pelabuhan",
                                  "3. Perkembangan Nilai Impor",
                                  "4. Perkembangan Nilai Impor (c-t-c)",
                                  "5. Volume Impor Menurut Pelabuhan"
                                ),
                                "Impor Menurut Negara Asal" = c(
                                  "6. Nilai Impor Menurut Negara Asal",
                                  "7. Nilai Impor Nonmigas Menurut Negara Asal",
                                  "8. Nilai Impor Migas Menurut Negara Asal",
                                  "9. Nilai Impor Negara Asal Utama HS2 Digit",
                                  "10. Perkembangan Nilai Impor Negara Asal Utama",
                                  "11. Nilai Impor Kumulatif Menurut Negara Asal",
                                  "12. Perkembangan Nilai Impor Negara Asal Utama (c-t-c)"
                                ),
                                "Impor Menurut HS2 Digit" = c(
                                  "13. Nilai Impor Nonmigas Menurut Golongan Barang HS2 Digit",
                                  "14. Peningkatan/Penurunan Nilai Impor Nonmigas HS2 Digit (m-t-m)",
                                  "15. Share Nilai Impor Nonmigas HS2 Digit"
                                )
                                
                              ),
                              selected = "1. Nilai Impor Menurut Sektor"
                            )
                            ,
                            div(
                              style = "display: flex; flex-wrap: wrap; gap: 10px; justify-content: flex-start;",
                              downloadButton("downloadDataImpor", "Download Tabel", style = "flex: 1 1 200px;"),
                              downloadButton("downloadAllImpor", "Download Semua Tabel", style = "flex: 1 1 200px;")
                            ),
                            tags$br(),
                            DTOutput("tabel_output_impor_3")
                          )
                          ,
                          
                          # Tab 2: Tabel Output
                          tabPanel(
                            title = "Data",
                            DTOutput("tabel_data_impor")
                          )
                        )
                      )
                      
                    )
                    
                  )),
          
          
          tabItem("neraca",
                  div(
                    id = "dashboard-content",
                    div(
                      style = "width: 100%; background-color: white; padding: 8px 12px; border-radius: 0 0 8px 8px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); margin-bottom: 15px;",
                      tags$h4(
                        icon("balance-scale"), 
                        HTML("&nbsp;Olah Data Neraca Perdagangan"),
                        style = "margin: 0;"
                      )
                    ),
                    
                    fluidRow(
                      box(
                        title = tagList(icon("upload"), HTML("&nbsp;&nbsp;Input Data Ekspor & Impor")),
                        status = "primary",
                        solidHeader = TRUE,
                        collapsible= FALSE,
                        width = 3,
                        fileInput("file_input_neraca_ekspor", "Pilih Data Ekspor (.sav)", accept = ".sav"),
                        fileInput("file_input_neraca_impor", "Pilih Data Impor (.sav)", accept = ".sav"),
                        selectInput("tahun_neraca", "Pilih Tahun", choices = NULL),
                        selectInput("bulan_neraca", "Pilih Bulan", choices = NULL),
                        tags$div(
                          style = "width: 100%;",
                          actionButton(
                            "analisis_button_neraca", 
                            label = tagList(
                              icon("magnifying-glass-chart", style = "color: white;"),
                              tags$span("Analisis", style = "color: white;")
                            ), 
                            class = "btn btn-primary", 
                            style = "width: 100%;"
                          )
                        )
                        
                      ),
                      
                      box(
                        title = tagList(icon("table"), HTML("&nbsp;&nbsp;Tabel Output Neraca Perdagangan")),
                        status = "primary",
                        solidHeader = TRUE,
                        collapsible= FALSE,
                        width = 9,
                        
                        tabBox(
                          width = 12,
                          
                          # Tab 1: Pilih Tabel + Tombol Download
                          tabPanel(
                            title = "Tabel Output",
                            selectInput(
                              inputId = "pilih_tabel_neraca",
                              label = NULL,
                              choices = list(
                                "1. Nilai Neraca Perdagangan",
                                "2. Perkembangan Nilai Neraca Perdagangan",
                                "3. Perkembangan Nilai Neraca Perdagangan (c-t-c)"
                              ),
                              selected = "1. Nilai Neraca Perdagangan"
                            )
                            ,
                            div(
                              style = "display: flex; flex-wrap: wrap; gap: 10px; justify-content: flex-start;",
                              downloadButton("downloadDataNeraca", "Download Tabel", style = "flex: 1 1 200px;"),
                              downloadButton("downloadAllNeraca", "Download Semua Tabel", style = "flex: 1 1 200px;")
                            ),
                            tags$br(),
                            DTOutput("tabel_output_neraca_3")
                          )
                          ,
                          
                          # Tab 2: Tabel Output
                          tabPanel(
                            title = "Data Ekspor",
                            DTOutput("tabel_data_neraca_ekspor")
                          ),
                          
                          tabPanel(
                            title = "Data Impor",
                            DTOutput("tabel_data_neraca_impor")
                          )
                        )
                      )
                      
                    )
                    
                  )),
          
          tabItem(
            tabName = "calendar_events",
            div(
              id = "dashboard-content",
              div(
                style = "width: 100%; background-color: white; padding: 8px 12px; border-radius: 0 0 8px 8px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); margin-bottom: 15px;",
                tags$h4(
                  icon("calendar-alt"), 
                  HTML("&nbsp;Kalender Kegiatan"),
                  style = "margin: 0;"
                )
              ),
            fluidRow(
              box(
                title = tagList(icon("calendar-alt"), HTML("&nbsp;&nbsp;Kalender Publikasi Ekspor Impor")),
                width = 12,
                status = "primary",
                solidHeader = TRUE,
                collapsible = FALSE,
                tags$div(
                  style = "width: 100%; height: 450px;",  # Sesuaikan tinggi di sini jika perlu
                  tags$iframe(
                    src = "https://calendar.google.com/calendar/embed?hl=id&src=b57de73f82e11e8c1bb19baec1bbe66772f98ddca294b76444e8cfbdadbaa6bb%40group.calendar.google.com&src=id.indonesian%23holiday%40group.v.calendar.google.com&color=%23007bff&color=%23dc3545&ctz=Asia%2FJakarta&mode=MONTH&showPrint=0&showTabs=0&showCalendars=0&showTz=0&showTitle=0",
                    style = "width: 100%; height: 100%; border: 0;",
                    frameborder = "0",
                    scrolling = "no"
                  )
                )
              )
            )
          )),
          
          
          tabItem("faq",
                  div(
                    id = "dashboard-content",
                    div(
                      style = "width: 100%; background-color: white; padding: 8px 12px; border-radius: 0 0 8px 8px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); margin-bottom: 15px;",
                      tags$h4(
                        icon("comments"), 
                        HTML("&nbsp;FAQ & Saran"),
                        style = "margin: 0;"
                      )
                    ),
                    
                    fluidRow(
                      # Kolom kiri: Daftar Pertanyaan FAQ
                      box(
                        title = tagList(icon("question-circle"), HTML("&nbsp;&nbsp;Frequently Ask Question")),
                        status = "primary",
                        solidHeader = TRUE,
                        collapsible= FALSE,
                        width = 7,
                        
                        box(
                          title = tagList(icon("users"), HTML("&nbsp;&nbsp;Siapa saja yang dapat mengakses SIPRO?")),
                          status = "primary",
                          solidHeader = TRUE,
                          collapsible = TRUE,
                          collapsed = FALSE,
                          width = 12,
                          p("SIPRO dapat diakses oleh petugas pengolah data ekspor-impor di BPS Kabupaten Karimun, khususnya yang memiliki kode akses sistem.")
                        ),
                        
                        box(
                          title = tagList(icon("exchange-alt"), HTML("&nbsp;&nbsp;Seperti apa konsep Ekspor dan Impor yang digunakan?")),
                          status = "primary",
                          solidHeader = TRUE,
                          collapsible = TRUE,
                          collapsed = TRUE,
                          width = 12,
                          tags$p(HTML("Konsep Ekspor dan Impor yang digunakan dalam aplikasi ini mengacu pada seluruh barang yang keluar dari atau masuk ke wilayah Indonesia melalui pelabuhan yang berada di Kabupaten Karimun, tanpa mempertimbangkan asal atau tujuan akhir barang tersebut.")),
                          tags$p(HTML("Sebagai contoh, barang yang berasal dari daerah lain di Indonesia namun dikirim ke luar negeri melalui pelabuhan di Kabupaten Karimun tetap dikategorikan sebagai <em>ekspor Kabupaten Karimun</em>. Sebaliknya, barang dari luar negeri yang masuk ke Indonesia melalui pelabuhan di Kabupaten Karimun dikategorikan sebagai <em>impor Kabupaten Karimun</em>, meskipun tujuan akhirnya bukan Kabupaten Karimun."))
                        )
                        ,
                        
                        box(
                          title = tagList(icon("dollar-sign"), HTML("&nbsp;&nbsp;Satuan apa yang digunakan pada Nilai dan Volume?")),
                          status = "primary",
                          solidHeader = TRUE,
                          collapsible = TRUE,
                          collapsed = TRUE,
                          width = 12,
                          tags$p(HTML("Satuan yang digunakan untuk Nilai adalah <em>Juta US Dollar (US$)</em>, sedangkan satuan untuk Volume adalah <em>Ribu Ton</em>."))
                        )
                        ,
                        
                        box(
                          title = tagList(icon("undo-alt"), HTML("&nbsp;&nbsp;Mengapa terdapat Indonesia sebagai Negara Asal pada tabel Impor?")),
                          status = "primary",
                          solidHeader = TRUE,
                          collapsible = TRUE,
                          collapsed = TRUE,
                          width = 12,
                          tags$p(HTML("Hal tersebut terjadi karena barang tersebut merupakan barang <em>reimpor</em>, yaitu barang asal Indonesia yang sebelumnya diekspor ke luar negeri dan kemudian diimpor kembali ke Indonesia.")),
                          tags$p(HTML("Reimpor dapat terjadi karena berbagai alasan, seperti barang yang ditolak oleh negara tujuan, adanya kerusakan, atau kebutuhan untuk diperbaiki. Meskipun berasal dari Indonesia, secara prosedur dan pencatatan, barang tersebut tetap dicatat sebagai impor karena masuk kembali dari luar negeri."))
                  
                        )
                        ,
                        
                        box(
                          title = tagList(icon("exclamation-circle"), HTML("&nbsp;&nbsp;Apa bedanya \"-\" dan \"0,00\" pada tabel output?")),
                          status = "primary",
                          solidHeader = TRUE,
                          collapsible = TRUE,
                          collapsed = TRUE,
                          width = 12,
                          p("\"-\" menunjukkan bahwa data pada periode tersebut tidak tersedia atau bernilai kosong, sehingga tidak memungkinkan dilakukan operasi pembagian untuk menghitung pertumbuhan (m-to-m, y-on-y, atau c-to-c). Sementara \"0,00\" memiliki arti bahwa nilainya ada namun sangat kecil sehingga ketika dibulatkan menjadi bilangan desimal dengan 2 digit angka di belakang koma menjadi 0,00.")
                        )
                        
                      )
                      ,
                      
                      # Kolom kanan: Formulir Masukan & Saran
                      box(
                        title = tagList(icon("comment-dots"), HTML("&nbsp;&nbsp;Kirim Masukan & Saran")),
                        status = "primary",
                        solidHeader = TRUE,
                        collapsible = FALSE,
                        width = 5,
                        useShinyjs(),
                        
                        textInput(
                          inputId = "namaPengguna", 
                          label = "Nama Anda", 
                          placeholder = "Masukkan nama (minimal 1 kata)"
                        ),
                        
                        textAreaInput(
                          inputId = "masukanPengguna", 
                          label = "Masukan atau Saran", 
                          placeholder = "Tulis saran atau masukan Anda di sini (minimal 5 kata)", 
                          rows = 5
                        ),
                        
                        tags$div(
                          style = "width: 100%;",
                          actionButton(
                            inputId = "kirimSaran",
                            label = tagList(icon("paper-plane", style = "color: white;"),
                                            tags$span("Kirim", style = "color: white;")),
                            class = "btn btn-primary",
                            style = "width: 100%;",
                            disabled = TRUE
                          )
                        ),
                        
                        br(),
                        
                        actionButton("lihatSaran", label = "Lihat Saran", class = "btn btn-secondary", style = "width: 100%;"),
                        
                        br(), br(),
                        
                        uiOutput("saranTerimaKasih")
                      )
                      
                      
                    ),
                    
                    # Tombol WhatsApp Mengambang
                    tags$head(
                      tags$style(HTML("
      .wa-float {
        position: fixed;
        bottom: 70px;
        right: 25px;
        background-color: #25D366;
        color: white;
        border-radius: 30px;
        padding: 10px 16px;
        font-size: 16px;
        display: flex;
        align-items: center;
        gap: 8px;
        box-shadow: 2px 2px 5px rgba(0,0,0,0.3);
        text-decoration: none;
        z-index: 9999;
        transition: background-color 0.3s;
      }
      .wa-float:hover {
        background-color: #1ebd5a;
        text-decoration: none;
        color: white;
      }
      .wa-icon {
        font-size: 22px;
      }
    "))
                    ),
                    
                    tags$a(
                      href = "https://wa.me/6285263715056?text=Halo%2C%20saya%20ingin%20bertanya%20terkait%20SIPRO",
                      target = "_blank",
                      class = "wa-float",
                      HTML('<i class="fab fa-whatsapp wa-icon"></i> Contact Person')
                    )
                    
                  )
                  
                  
          ))),
    br(),br(),
    tags$footer(class = "main-footer",
                tags$div(style = "
  height: 5px;
  background-image: repeating-linear-gradient(to right,
    #3c8dbc,
    #3c8dbc 33.33%,
    #f39c12 33.33%,
    #f39c12 66.66%,
    #00a65a 66.66%,
    #00a65a 100%
  );
"),
                tags$strong("Copyright "),
                icon("copyright", style = "font-size: 12px; margin: 0 4px;"),
                tags$strong("BPS KABUPATEN KARIMUN", style = "color: #3c8dbc;")
    )
  )
)


server <- function(input, output, session) {
  
  
  
  data_tabel <- data.frame(
    'Tabel' = 1:15,
    'Ekspor' = c(
      "Nilai Ekspor Menurut Sektor",
      "Nilai Ekspor Menurut Pelabuhan",
      "Perkembangan Nilai Ekspor",
      "Perkembangan Nilai Ekspor (c-t-c)",
      "Volume Ekspor Menurut Pelabuhan",
      "Nilai Ekspor Menurut Negara Tujuan",
      "Nilai Ekspor Nonmigas Menurut Negara Tujuan",
      "Nilai Ekspor Migas Menurut Negara Tujuan",
      "Nilai Ekspor Negara Tujuan Utama HS2 Digit",
      "Perkembangan Nilai Ekspor Negara Tujuan Utama",
      "Nilai Ekspor Kumulatif Menurut Negara Tujuan",
      "Perkembangan Nilai Ekspor Negara Tujuan Utama (c-t-c)",
      "Nilai Ekspor Nonmigas Menurut Golongan Barang HS2 Digit",
      "Peningkatan/Penurunan Nilai Ekspor Nonmigas HS2 Digit (m-t-m)",
      "Share Nilai Ekspor Nonmigas HS2 Digit"
    ),
    'Impor' = c(
      "Nilai Impor Menurut Sektor",
      "Nilai Impor Menurut Pelabuhan",
      "Perkembangan Nilai Impor",
      "Perkembangan Nilai Impor (c-t-c)",
      "Volume Impor Menurut Pelabuhan",
      "Nilai Impor Menurut Negara Asal",
      "Nilai Impor Nonmigas Menurut Negara Asal",
      "Nilai Impor Migas Menurut Negara Asal",
      "Nilai Impor Negara Asal Utama HS2 Digit",
      "Perkembangan Nilai Impor Negara Asal Utama",
      "Nilai Impor Kumulatif Menurut Negara Asal",
      "Perkembangan Nilai Impor Negara Asal Utama (c-t-c)",
      "Nilai Impor Nonmigas Menurut Golongan Barang HS2 Digit",
      "Peningkatan/Penurunan Nilai Impor Nonmigas HS2 Digit (m-t-m)",
      "Share Nilai Impor Nonmigas HS2 Digit"
      
    ),
    'Neraca Perdagangan' = c(
      "Nilai Neraca Perdagangan",
      "Perkembangan Nilai Neraca Perdagangan",
      "Perkembangan Nilai Neraca Perdagangan (c-t-c)",
      "-",
      "-",
      "-",
      "-",
      "-",
      "-",
      "-",
      "-",
      "-",
      "-",
      "-",
      "-"
    ),
    check.names = FALSE
  )
  output$tabelOutput <- renderDT({
    datatable(data_tabel, rownames = FALSE, options = list(
      scrollX = TRUE,         # Menambahkan scroll horizontal
      scrollY = "300px",      # Menambahkan scroll vertikal dengan tinggi 400px
      fixedHeader = TRUE,     # Membuat header tetap (sticky)
      ordering = FALSE        # Menonaktifkan fitur sorting otomatis
    ))
  })
  
  
  # password yang benar
  kode_benar <- "bps2101"
  
  # reactive untuk menyimpan akses
  akses_diberikan <- reactiveVal(FALSE)
  alert_sedang_tampil <- reactiveVal(FALSE)
  
  
  # Fungsi untuk meminta password
  tampilkan_kode_akses <- function(){
    if (!alert_sedang_tampil()) {
      alert_sedang_tampil(TRUE)
      # Blur
      shinyjs::addClass("app_content", "blurred")
      shinyjs::runjs("document.querySelector('.main-sidebar').classList.add('blurred');")
      shinyjs::runjs("document.querySelector('.main-footer').classList.add('blurred');")
      shinyjs::runjs("document.querySelector('.main-header.navbar').classList.add('blurred');")
      
      # Tampilkan alert password
      shinyalert(
        title = "Kode Akses",
        text = "Masukkan kode akses untuk masuk ke aplikasi.",
        type = "input",
        inputType = "password",
        closeOnEsc = FALSE,
        closeOnClickOutside = FALSE,
        showCancelButton = FALSE,
        animation = TRUE,
        callbackR = function(value) {
          if (value == kode_benar) {
            akses_diberikan(TRUE)
            # Hilangkan blur
            shinyjs::removeClass("app_content", "blurred")
            shinyjs::runjs("document.querySelector('.main-sidebar').classList.remove('blurred');")
            shinyjs::runjs("document.querySelector('.main-footer').classList.remove('blurred');")
            shinyjs::runjs("document.querySelector('.main-header.navbar').classList.remove('blurred');")
            
            # Simpan di localStorage
            expiry_time <- as.numeric(Sys.time()) + 60 * 60
            shinyjs::runjs(sprintf(
              "localStorage.setItem('akses_diberikan', 'true'); localStorage.setItem('akses_expiry', '%f');",
              expiry_time
            ))
          } else {
            # Coba lagi
            alert_sedang_tampil(FALSE)
            shinyalert("Kode Salah!", "Silakan coba lagi.", 
                       type = "error",
                       closeOnEsc = TRUE,
                       closeOnClickOutside = TRUE,
                       callbackR = function(x) {
                         tampilkan_kode_akses()
                       })
          }
        }
      )
    }
  }
  
  # Cek akses di localStorage saat app dimulai
  observe({ 
    # perintah javascript untuk mengambil dari localStorage
    shinyjs::runjs("
    (function(){
      var akses = localStorage.getItem('akses_diberikan');
      var expiry = localStorage.getItem('akses_expiry'); 
      var now = Math.floor(Date.now()/1000);
      if (akses === 'true' && expiry && parseInt(expiry) > now) {
        Shiny.setInputValue('akses_sudah', true);
      } else {
        localStorage.removeItem('akses_diberikan');
        localStorage.removeItem('akses_expiry');
        Shiny.setInputValue('akses_sudah', false);
      }
    })()
  ") 
  })
  
  
  # Saat diberitahu oleh browser
  observeEvent(input$akses_sudah, {
    req(!is.null(input$akses_sudah))
    
    if (isTRUE(input$akses_sudah)) {
      akses_diberikan(TRUE)
    } else {
      akses_diberikan(FALSE)
      tampilkan_kode_akses()
    }
  })
  
  
  # Jika memang false dan modal juga tampil
  observe({ 
    req(input$akses_sudah == FALSE)
    
    if (!akses_diberikan() && !alert_sedang_tampil()) {
      tampilkan_kode_akses()
    }
  })
  
  
  
  
  
  

  
  
  ##---------------------------------------------------------Ekspor-------------------------------------------
  
  # Saat file di-upload, update pilihan tahun
  # Reactive untuk menyimpan data yang di-upload
  data_reaktif <- reactiveVal(NULL)
  
  # Saat file di-upload, baca datanya dan update pilihan tahun
  observeEvent(input$file_input, {
    req(input$file_input)
    
    df <- as.data.frame(read_sav(input$file_input$datapath))
    
    # Cek apakah kolom TAHUN dan BULAN ada
    if (any(c("TAHUN", "BULAN", "NILAI", "BERAT", "PELABUHAN", 
              "HS", "NEGARA", "JENIS", "SEKTOR", "KOMODITI", 
              "HS2", "HSDUA", "NEGARA2", "PELABUHAN1") %in% names(df) == FALSE)) {
      showNotification("Ada kolom penting yang tidak ditemukan di data ekspor", type = "error")
      return()
    }
    
    
    data_reaktif(df)  # Simpan ke reactive
    
    # Update selectInput tahun
    updateSelectInput(session, "tahun", choices = sort(unique(df$TAHUN)))
    
    # Reset bulan saat file baru diupload
    updateSelectInput(session, "bulan", choices = NULL)
  })
  
  # Saat tahun dipilih, update bulan
  observeEvent(input$tahun, {
    req(data_reaktif())
    req(input$tahun)
    
    df <- data_reaktif()
    
    bulan_angka <- df %>%
      filter(TAHUN == input$tahun) %>%
      pull(BULAN) %>%
      as.numeric() %>%     # konversi ke angka
      unique() %>%
      sort()               # sort sebagai angka (1–12)
    
    # Filter hanya bulan yang valid
    bulan_angka <- bulan_angka[!is.na(bulan_angka) & bulan_angka %in% 1:12]
    
    # Konversi ke nama bulan
    nama_bulan <- month.name[bulan_angka]
    names(bulan_angka) <- nama_bulan
    
    # Update selectInput
    updateSelectInput(session, "bulan", choices = bulan_angka)
    
    
    updateSelectInput(session, "bulan", choices = bulan_angka)
  })
  
  
  observeEvent(input$analisis_button, {
    
    req(data_reaktif())
    req(input$tahun)
    req(input$bulan)
    
    bulan<-as.numeric(input$bulan)
    tahun <- as.numeric(input$tahun) 
    data <- data_reaktif()
    
    
    data$BULAN<-as.numeric(data$BULAN)
    data$TAHUN<-as.numeric(data$TAHUN)
    
    data$PROPINSI<-as.character(data$PROPINSI)
    data$PROPINSI <- trimws(data$PROPINSI, which = "both")
    
    data$PELABUHAN<-as.character(data$PELABUHAN)
    data$PELABUHAN <- trimws(data$PELABUHAN, which = "both")
    
    data$HS<-as.character(data$HS)
    data$HS<-ifelse(nchar(data$HS)==7, paste0("0",data$HS), data$HS)
    data$HS <- trimws(data$HS, which = "both")
    
    data$NEGARA<-as.character(data$NEGARA)
    data$NEGARA <- trimws(data$NEGARA, which = "both")
    
    data$BERAT<-as.numeric(data$BERAT)
    data$NILAI<-as.numeric(data$NILAI)
    
    
    data$HSDUA<-as.character(data$HSDUA)
    data$HSDUA<-ifelse(nchar(data$HSDUA)==1, paste0("0",data$HSDUA), data$HSDUA)
    data$HSDUA <- trimws(data$HSDUA, which = "both")
    
    data$KABKOT<-as.character(data$KABKOT)
    data$KABKOT <- trimws(data$KABKOT, which = "both")
    
    data$PELABUHAN1<-as.character(data$PELABUHAN1)
    data$PELABUHAN1 <- trimws(data$PELABUHAN1, which = "both")
    
    data$NEGARA2<-as.character(data$NEGARA2)
    data$NEGARA2 <- trimws(data$NEGARA2, which = "both")
    
    data$HS2<-as.character(data$HS2)
    data$HS2 <- trimws(data$HS2, which = "both")
    
    data$JENIS<-as.character(data$JENIS)
    data$JENIS <- trimws(data$JENIS, which = "both")
    
    data$SEKTOR<-as.character(data$SEKTOR)
    data$SEKTOR <- trimws(data$SEKTOR, which = "both")
    
    data$KOMODITI<-as.character(data$KOMODITI)
    data$KOMODITI <- trimws(data$KOMODITI, which = "both")
    
    data<-mutate(data,NILAI_JUTA=NILAI/1000000)
    ##data <- data[data$PROVASAL == "21", ]
    data$NILAI_JUTA<-as.numeric(data$NILAI_JUTA)
    
    
    ## ----------------------------------------------------------------- SLIDE 1 ---------------------------------------------------------------------------------------------
    ##Ekspor Bulan Ini
    ekspor_bulan_ini<-sum(data$NILAI_JUTA[data$BULAN==bulan & data$TAHUN==tahun])
    ekspor_bulan_ini_migas<-sum(data$NILAI_JUTA[data$JENIS=="MIGAS" & data$BULAN==bulan & data$TAHUN==tahun])
    ekspor_bulan_ini_nonmigas<-sum(data$NILAI_JUTA[data$JENIS=="NON MIGAS" & data$BULAN==bulan & data$TAHUN==tahun])
    kolom_bulan_ini<-paste(month.abb[bulan]," ", tahun)
    
    ##Ekspor Bulan Sebelumnya (M to M)
    bulan_sebelumnya <- ifelse(bulan-1 == 0, 12, bulan-1)
    tahun_sebelumnya <- ifelse(bulan_sebelumnya == 12, tahun-1, tahun)
    
    ekspor_bulan_lalu<-sum(data$NILAI_JUTA[data$BULAN==bulan_sebelumnya & data$TAHUN==tahun_sebelumnya])
    ekspor_bulan_lalu_migas<-sum(data$NILAI_JUTA[data$JENIS=="MIGAS" & data$BULAN==bulan_sebelumnya & data$TAHUN==tahun_sebelumnya])
    ekspor_bulan_lalu_nonmigas<-sum(data$NILAI_JUTA[data$JENIS=="NON MIGAS" & data$BULAN==bulan_sebelumnya & data$TAHUN==tahun_sebelumnya])
    
    kolom_bulan_lalu<-paste(month.abb[bulan_sebelumnya]," ", tahun_sebelumnya)
    
    ##Ekspor Tahun Sebelumnya (Y on Y)
    ekspor_tahun_lalu<-sum(data$NILAI_JUTA[data$BULAN==bulan & data$TAHUN==tahun-1])
    ekspor_tahun_lalu_migas<-sum(data$NILAI_JUTA[data$JENIS=="MIGAS" & data$BULAN==bulan & data$TAHUN==tahun-1])
    ekspor_tahun_lalu_nonmigas<-sum(data$NILAI_JUTA[data$JENIS=="NON MIGAS" & data$BULAN==bulan & data$TAHUN==tahun-1])
    
    kolom_tahun_lalu<-paste(month.abb[bulan]," ", tahun-1)
    
    
    ##Ekspor Cumulative (C to C)
    ekspor_cumulative<-sum(data$NILAI_JUTA[data$BULAN<=bulan & data$TAHUN==tahun])
    ekspor_cumulative_migas<-sum(data$NILAI_JUTA[data$JENIS=="MIGAS" & data$BULAN<=bulan & data$TAHUN==tahun])
    ekspor_cumulative_nonmigas<-sum(data$NILAI_JUTA[data$JENIS=="NON MIGAS" & data$BULAN<=bulan & data$TAHUN==tahun])
    
    ekspor_cumulative_tahun_lalu<-sum(data$NILAI_JUTA[data$BULAN<=bulan & data$TAHUN==tahun-1])
    ekspor_cumulative_migas_tahun_lalu<-sum(data$NILAI_JUTA[data$JENIS=="MIGAS" & data$BULAN<=bulan & data$TAHUN==tahun-1])
    ekspor_cumulative_nonmigas_tahun_lalu<-sum(data$NILAI_JUTA[data$JENIS=="NON MIGAS" & data$BULAN<=bulan & data$TAHUN==tahun-1])
    
    kolom_cumulative<-paste(month.abb[1],"-", month.abb[bulan]," ", tahun)
    kolom_cumulative_tahun_lalu<-paste(month.abb[1],"-", month.abb[bulan]," ", tahun-1)
    
    
    ##Tabel Output
    sektor_cumulative_tahun_lalu<-summarise(group_by(data[data$BULAN<=bulan & data$TAHUN==tahun-1,], JENIS,SEKTOR), Total_Ekspor = sum(NILAI_JUTA))
    sektor_cumulative<-summarise(group_by(data[data$BULAN<=bulan & data$TAHUN==tahun,], JENIS,SEKTOR), Total_Ekspor = sum(NILAI_JUTA))
    sektor_bulan_ini<-summarise(group_by(data[data$BULAN==bulan & data$TAHUN==tahun,], JENIS,SEKTOR), Total_Ekspor = sum(NILAI_JUTA))
    sektor_bulan_lalu<-summarise(group_by(data[data$BULAN==bulan_sebelumnya & data$TAHUN==tahun_sebelumnya,], JENIS,SEKTOR), Total_Ekspor = sum(NILAI_JUTA))
    sektor_tahun_lalu<-summarise(group_by(data[data$BULAN==bulan & data$TAHUN==tahun-1,], JENIS,SEKTOR), Total_Ekspor = sum(NILAI_JUTA))
    
    
    if (bulan!=1) {
      gabung1<-full_join(sektor_bulan_ini,sektor_cumulative,by=c("JENIS","SEKTOR"))
      gabung2<-full_join(sektor_bulan_lalu,gabung1,by=c("JENIS","SEKTOR"))
      gabung3<-full_join(sektor_cumulative_tahun_lalu,gabung2,by=c("JENIS","SEKTOR"))
      tabel1<-full_join(sektor_tahun_lalu,gabung3,by=c("JENIS","SEKTOR"))
      colnames(tabel1)<-c("Jenis","Uraian",kolom_tahun_lalu, kolom_cumulative_tahun_lalu,kolom_bulan_lalu, kolom_bulan_ini, kolom_cumulative)
      tabel1 <- arrange(tabel1,  desc(!!sym(kolom_cumulative)),desc(!!sym(kolom_cumulative_tahun_lalu)))
      
      tabel1_gabung<-data.frame(Jenis=c("","",""),Uraian=c("TOTAL EKSPOR", "MIGAS", "NON MIGAS"),kolom_tahun_lalu=c(ekspor_tahun_lalu,ekspor_tahun_lalu_migas,ekspor_tahun_lalu_nonmigas), kolom_cumulative_tahun_lalu=c(ekspor_cumulative_tahun_lalu,ekspor_cumulative_migas_tahun_lalu,ekspor_cumulative_nonmigas_tahun_lalu),kolom_bulan_lalu=c(ekspor_bulan_lalu,ekspor_bulan_lalu_migas,ekspor_bulan_lalu_nonmigas),kolom_bulan_ini=c(ekspor_bulan_ini,ekspor_bulan_ini_migas,ekspor_bulan_ini_nonmigas), kolom_cumulative=c(ekspor_cumulative,ekspor_cumulative_migas,ekspor_cumulative_nonmigas)) 
      names(tabel1_gabung) <- names(tabel1)
      tabel1<-rbind(tabel1_gabung,tabel1)
      
    } else {
      gabung1<-full_join(sektor_bulan_lalu,sektor_bulan_ini,by=c("JENIS","SEKTOR"))
      tabel1<-full_join(sektor_tahun_lalu,gabung1,by=c("JENIS","SEKTOR"))
      colnames(tabel1)<-c("Jenis","Uraian",kolom_tahun_lalu, kolom_bulan_lalu, kolom_bulan_ini)
      tabel1 <- arrange(tabel1, desc(!!sym(kolom_bulan_ini)), desc(!!sym(kolom_bulan_lalu)))
      
      tabel1_gabung<-data.frame(Jenis=c("","",""),Uraian=c("TOTAL EKSPOR", "MIGAS", "NON MIGAS"),kolom_tahun_lalu=c(ekspor_tahun_lalu,ekspor_tahun_lalu_migas,ekspor_tahun_lalu_nonmigas), kolom_bulan_lalu=c(ekspor_bulan_lalu,ekspor_bulan_lalu_migas,ekspor_bulan_lalu_nonmigas),kolom_bulan_ini=c(ekspor_bulan_ini,ekspor_bulan_ini_migas,ekspor_bulan_ini_nonmigas))      
      names(tabel1_gabung) <- names(tabel1)
      tabel1<-rbind(tabel1_gabung,tabel1)
    }
    
    tabel1$urutan <- NA
    
    # Urutan utama
    tabel1$urutan[tabel1$Uraian == "TOTAL EKSPOR"] <- 1
    tabel1$urutan[tabel1$Uraian == "MIGAS"] <- 2
    
    # Urut sektor MIGAS
    i <- 3
    for (j in 1:nrow(tabel1)) {
      if (tabel1$Jenis[j] == "MIGAS" & !(tabel1$Uraian[j] %in% c("TOTAL EKSPOR", "MIGAS", "NON MIGAS"))) {
        tabel1$urutan[j] <- i
        i <- i + 1
      }
    }
    
    # Urutan NON MIGAS
    tabel1$urutan[tabel1$Uraian == "NON MIGAS"] <- i
    i <- i + 1
    
    # Urut sektor NON MIGAS
    for (j in 1:nrow(tabel1)) {
      if (tabel1$Jenis[j] == "NON MIGAS" & !(tabel1$Uraian[j] %in% c("TOTAL EKSPOR", "MIGAS", "NON MIGAS"))) {
        tabel1$urutan[j] <- i
        i <- i + 1
      }
    }
    
    # Urutkan
    tabel1 <- tabel1[order(tabel1$urutan), ]
    tabel1$Uraian <- ifelse(!(tabel1$Uraian %in% c("TOTAL EKSPOR","MIGAS", "NON MIGAS")), paste(" - ", tabel1$Uraian), tabel1$Uraian)
    tabel1$urutan<-NULL
    tabel1$Jenis<-NULL
    tabel1[tabel1 == 0] <- NA
    
    if(bulan!=1){
      tabel1<-mutate(tabel1, tabel1_mtm=100*(!!sym(kolom_bulan_ini)-!!sym(kolom_bulan_lalu))/!!sym(kolom_bulan_lalu), tabel1_yoy=100*(!!sym(kolom_bulan_ini)-!!sym(kolom_tahun_lalu))/!!sym(kolom_tahun_lalu), tabel1_ctc=100*(!!sym(kolom_cumulative)-!!sym(kolom_cumulative_tahun_lalu))/!!sym(kolom_cumulative_tahun_lalu), peran=100*!!sym(kolom_cumulative)/ekspor_cumulative)
      colnames(tabel1)<-c("Uraian",kolom_tahun_lalu, kolom_cumulative_tahun_lalu,kolom_bulan_lalu, kolom_bulan_ini, kolom_cumulative, "m-to-m (%)", "y-on-y (%)", "c-to-c (%)", paste("Peran (%)",kolom_cumulative))
    } else{
      tabel1<-mutate(tabel1, tabel1_mtm=100*(!!sym(kolom_bulan_ini)-!!sym(kolom_bulan_lalu))/!!sym(kolom_bulan_lalu), tabel1_yoy=100*(!!sym(kolom_bulan_ini)-!!sym(kolom_tahun_lalu))/!!sym(kolom_tahun_lalu),peran=100*!!sym(kolom_bulan_ini)/ekspor_bulan_ini)
      colnames(tabel1)<-c("Uraian",kolom_tahun_lalu,kolom_bulan_lalu, kolom_bulan_ini, "m-to-m (%)", "y-on-y (%)", paste("Peran (%)",kolom_bulan_ini))
    }
    
    tabel1 <- tabel1 %>%
      mutate(
        `m-to-m (%)` = ifelse(!is.na(!!sym(kolom_bulan_lalu)) & is.na(!!sym(kolom_bulan_ini)), -100, `m-to-m (%)`),
        `y-on-y (%)` = ifelse(!is.na(!!sym(kolom_tahun_lalu)) & is.na(!!sym(kolom_bulan_ini)), -100, `y-on-y (%)`)
      )
    
    if(bulan != 1){
      tabel1 <- tabel1 %>%
        mutate(
          `c-to-c (%)` = ifelse(!is.na(!!sym(kolom_cumulative_tahun_lalu)) & is.na(!!sym(kolom_cumulative)), -100, `c-to-c (%)`)
        )
    }
    
    tabel1<<-tabel1
    
    
    
    ##Tabel Output Nilai Ekspor Nonmigas Menurut Golongan Barang HS2 Digit (Ekspor Menurut HS2 Digit)
    hs2_cumulative_tahun_lalu<-summarise(group_by(data[data$BULAN<=bulan & data$TAHUN==tahun-1 & data$JENIS=="NON MIGAS",], HSDUA,HS2), Total_Ekspor = sum(NILAI_JUTA))
    hs2_cumulative<-summarise(group_by(data[data$BULAN<=bulan & data$TAHUN==tahun & data$JENIS=="NON MIGAS",], HSDUA,HS2), Total_Ekspor = sum(NILAI_JUTA))
    hs2_bulan_ini<-summarise(group_by(data[data$BULAN==bulan & data$TAHUN==tahun & data$JENIS=="NON MIGAS",], HSDUA,HS2), Total_Ekspor = sum(NILAI_JUTA))
    hs2_bulan_lalu<-summarise(group_by(data[data$BULAN==bulan_sebelumnya & data$TAHUN==tahun_sebelumnya & data$JENIS=="NON MIGAS",], HSDUA,HS2), Total_Ekspor = sum(NILAI_JUTA))
    hs2_tahun_lalu<-summarise(group_by(data[data$BULAN==bulan & data$TAHUN==tahun-1 & data$JENIS=="NON MIGAS",], HSDUA,HS2), Total_Ekspor = sum(NILAI_JUTA))
    
    
    if (bulan!=1) {
      gabung1_hs2<-full_join(hs2_bulan_ini,hs2_cumulative,by=c("HSDUA","HS2"))
      gabung2_hs2<-full_join(hs2_bulan_lalu,gabung1_hs2,by=c("HSDUA","HS2"))
      gabung3_hs2<-full_join(hs2_cumulative_tahun_lalu,gabung2_hs2,by=c("HSDUA","HS2"))
      tabel1_hs2<-full_join(hs2_tahun_lalu,gabung3_hs2,by=c("HSDUA","HS2"))
      colnames(tabel1_hs2)<-c("HS2 Digit","Deskripsi",kolom_tahun_lalu, kolom_cumulative_tahun_lalu,kolom_bulan_lalu, kolom_bulan_ini, kolom_cumulative)
      tabel1_1 <- arrange(tabel1_hs2, desc(!!sym(kolom_cumulative)),desc(!!sym(kolom_cumulative_tahun_lalu)))
      
    } else {
      gabung1_hs2<-full_join(hs2_bulan_lalu,hs2_bulan_ini,by=c("HSDUA","HS2"))
      tabel1_1<-full_join(hs2_tahun_lalu,gabung1_hs2,by=c("HSDUA","HS2"))
      colnames(tabel1_1)<-c("HS2 Digit","Deskripsi",kolom_tahun_lalu, kolom_bulan_lalu, kolom_bulan_ini)
      tabel1_1 <- arrange(tabel1_1, desc(!!sym(kolom_bulan_ini)),desc(!!sym(kolom_bulan_lalu)))
    }
    
    
    if(bulan!=1){
      tabel1_1<-mutate(tabel1_1, tabel1_1_mtm=100*(!!sym(kolom_bulan_ini)-!!sym(kolom_bulan_lalu))/!!sym(kolom_bulan_lalu), tabel1_1_yoy=100*(!!sym(kolom_bulan_ini)-!!sym(kolom_tahun_lalu))/!!sym(kolom_tahun_lalu), tabel1_1_ctc=100*(!!sym(kolom_cumulative)-!!sym(kolom_cumulative_tahun_lalu))/!!sym(kolom_cumulative_tahun_lalu), peran=100*!!sym(kolom_cumulative)/ekspor_cumulative_nonmigas)
      colnames(tabel1_1)<-c("Kode HS","Deskripsi",kolom_tahun_lalu, kolom_cumulative_tahun_lalu,kolom_bulan_lalu, kolom_bulan_ini, kolom_cumulative, "m-to-m (%)", "y-on-y (%)", "c-to-c (%)", paste("Peran (%)",kolom_cumulative))
    } else{
      tabel1_1<-mutate(tabel1_1, tabel1_1_mtm=100*(!!sym(kolom_bulan_ini)-!!sym(kolom_bulan_lalu))/!!sym(kolom_bulan_lalu), tabel1_1_yoy=100*(!!sym(kolom_bulan_ini)-!!sym(kolom_tahun_lalu))/!!sym(kolom_tahun_lalu),peran=100*!!sym(kolom_bulan_ini)/ekspor_bulan_ini_nonmigas)
      colnames(tabel1_1)<-c("Kode HS","Deskripsi",kolom_tahun_lalu,kolom_bulan_lalu, kolom_bulan_ini, "m-to-m (%)", "y-on-y (%)", paste("Peran (%)",kolom_bulan_ini))
    }
    
    tabel1_1 <- tabel1_1 %>%
      mutate(
        `m-to-m (%)` = ifelse(!is.na(!!sym(kolom_bulan_lalu)) & is.na(!!sym(kolom_bulan_ini)), -100, `m-to-m (%)`),
        `y-on-y (%)` = ifelse(!is.na(!!sym(kolom_tahun_lalu)) & is.na(!!sym(kolom_bulan_ini)), -100, `y-on-y (%)`)
      )
    
    if(bulan != 1){
      tabel1_1 <- tabel1_1 %>%
        mutate(
          `c-to-c (%)` = ifelse(!is.na(!!sym(kolom_cumulative_tahun_lalu)) & is.na(!!sym(kolom_cumulative)), -100, `c-to-c (%)`)
        )
    }
    
    tabel13<<-tabel1_1
    
    
    
    ##Tabel Output Nilai Ekspor Nonmigas Menurut Negara Tujuan (Ekspor Menurut Negara Tujuan)
    negara_nonmigas_cumulative_tahun_lalu<-summarise(group_by(data[data$BULAN<=bulan & data$TAHUN==tahun-1 & data$JENIS=="NON MIGAS",], NEGARA2), Total_Ekspor = sum(NILAI_JUTA))
    negara_nonmigas_cumulative<-summarise(group_by(data[data$BULAN<=bulan & data$TAHUN==tahun & data$JENIS=="NON MIGAS",], NEGARA2), Total_Ekspor = sum(NILAI_JUTA))
    negara_nonmigas_bulan_ini<-summarise(group_by(data[data$BULAN==bulan & data$TAHUN==tahun & data$JENIS=="NON MIGAS",], NEGARA2), Total_Ekspor = sum(NILAI_JUTA))
    negara_nonmigas_bulan_lalu<-summarise(group_by(data[data$BULAN==bulan_sebelumnya & data$TAHUN==tahun_sebelumnya & data$JENIS=="NON MIGAS",], NEGARA2), Total_Ekspor = sum(NILAI_JUTA))
    negara_nonmigas_tahun_lalu<-summarise(group_by(data[data$BULAN==bulan & data$TAHUN==tahun-1 & data$JENIS=="NON MIGAS",], NEGARA2), Total_Ekspor = sum(NILAI_JUTA))
    
    
    if (bulan!=1) {
      gabung1_negara_nonmigas<-full_join(negara_nonmigas_bulan_ini,negara_nonmigas_cumulative,by="NEGARA2")
      gabung2_negara_nonmigas<-full_join(negara_nonmigas_bulan_lalu,gabung1_negara_nonmigas,by="NEGARA2")
      gabung3_negara_nonmigas<-full_join(negara_nonmigas_cumulative_tahun_lalu,gabung2_negara_nonmigas,by="NEGARA2")
      tabel1_negara_nonmigas<-full_join(negara_nonmigas_tahun_lalu,gabung3_negara_nonmigas,by="NEGARA2")
      colnames(tabel1_negara_nonmigas)<-c("Negara",kolom_tahun_lalu, kolom_cumulative_tahun_lalu,kolom_bulan_lalu, kolom_bulan_ini, kolom_cumulative)
      tabel1_2 <- arrange(tabel1_negara_nonmigas, desc(!!sym(kolom_cumulative)),  desc(!!sym(kolom_cumulative_tahun_lalu)))
      
    } else {
      gabung1_negara_nonmigas<-full_join(negara_nonmigas_bulan_lalu,negara_nonmigas_bulan_ini,by="NEGARA2")
      tabel1_2<-full_join(negara_nonmigas_tahun_lalu,gabung1_negara_nonmigas,by="NEGARA2")
      colnames(tabel1_2)<-c("Negara",kolom_tahun_lalu, kolom_bulan_lalu, kolom_bulan_ini)
      tabel1_2 <- arrange(tabel1_2, desc(!!sym(kolom_bulan_ini)), desc(!!sym(kolom_bulan_lalu)))
    }
    
    
    if(bulan!=1){
      tabel1_2<-mutate(tabel1_2, tabel1_2_mtm=100*(!!sym(kolom_bulan_ini)-!!sym(kolom_bulan_lalu))/!!sym(kolom_bulan_lalu), tabel1_2_yoy=100*(!!sym(kolom_bulan_ini)-!!sym(kolom_tahun_lalu))/!!sym(kolom_tahun_lalu), tabel1_2_ctc=100*(!!sym(kolom_cumulative)-!!sym(kolom_cumulative_tahun_lalu))/!!sym(kolom_cumulative_tahun_lalu), peran=100*!!sym(kolom_cumulative)/ekspor_cumulative_nonmigas)
      colnames(tabel1_2)<-c("Negara",kolom_tahun_lalu, kolom_cumulative_tahun_lalu,kolom_bulan_lalu, kolom_bulan_ini, kolom_cumulative, "m-to-m (%)", "y-on-y (%)", "c-to-c (%)", paste("Peran (%)",kolom_cumulative))
    } else{
      tabel1_2<-mutate(tabel1_2, tabel1_2_mtm=100*(!!sym(kolom_bulan_ini)-!!sym(kolom_bulan_lalu))/!!sym(kolom_bulan_lalu), tabel1_2_yoy=100*(!!sym(kolom_bulan_ini)-!!sym(kolom_tahun_lalu))/!!sym(kolom_tahun_lalu),peran=100*!!sym(kolom_bulan_ini)/ekspor_bulan_ini_nonmigas)
      colnames(tabel1_2)<-c("Negara",kolom_tahun_lalu,kolom_bulan_lalu, kolom_bulan_ini, "m-to-m (%)", "y-on-y (%)", paste("Peran (%)",kolom_bulan_ini))
    }
    
    tabel1_2 <- tabel1_2 %>%
      mutate(
        `m-to-m (%)` = ifelse(!is.na(!!sym(kolom_bulan_lalu)) & is.na(!!sym(kolom_bulan_ini)), -100, `m-to-m (%)`),
        `y-on-y (%)` = ifelse(!is.na(!!sym(kolom_tahun_lalu)) & is.na(!!sym(kolom_bulan_ini)), -100, `y-on-y (%)`)
      )
    
    if(bulan != 1){
      tabel1_2 <- tabel1_2 %>%
        mutate(
          `c-to-c (%)` = ifelse(!is.na(!!sym(kolom_cumulative_tahun_lalu)) & is.na(!!sym(kolom_cumulative)), -100, `c-to-c (%)`)
        )
    }
    
    tabel7<<-tabel1_2
    
    
    ##Tabel Output Nilai Ekspor Migas Menurut Negara Tujuan (Ekspor Menurut Negara Tujuan)
    negara_migas_cumulative_tahun_lalu<-summarise(group_by(data[data$BULAN<=bulan & data$TAHUN==tahun-1 & data$JENIS=="MIGAS",], NEGARA2), Total_Ekspor = sum(NILAI_JUTA))
    negara_migas_cumulative<-summarise(group_by(data[data$BULAN<=bulan & data$TAHUN==tahun & data$JENIS=="MIGAS",], NEGARA2), Total_Ekspor = sum(NILAI_JUTA))
    negara_migas_bulan_ini<-summarise(group_by(data[data$BULAN==bulan & data$TAHUN==tahun & data$JENIS=="MIGAS",], NEGARA2), Total_Ekspor = sum(NILAI_JUTA))
    negara_migas_bulan_lalu<-summarise(group_by(data[data$BULAN==bulan_sebelumnya & data$TAHUN==tahun_sebelumnya & data$JENIS=="MIGAS",], NEGARA2), Total_Ekspor = sum(NILAI_JUTA))
    negara_migas_tahun_lalu<-summarise(group_by(data[data$BULAN==bulan & data$TAHUN==tahun-1 & data$JENIS=="MIGAS",], NEGARA2), Total_Ekspor = sum(NILAI_JUTA))
    
    
    if (bulan!=1) {
      gabung1_negara_migas<-full_join(negara_migas_bulan_ini,negara_migas_cumulative,by="NEGARA2")
      gabung2_negara_migas<-full_join(negara_migas_bulan_lalu,gabung1_negara_migas,by="NEGARA2")
      gabung3_negara_migas<-full_join(negara_migas_cumulative_tahun_lalu,gabung2_negara_migas,by="NEGARA2")
      tabel1_negara_migas<-full_join(negara_migas_tahun_lalu,gabung3_negara_migas,by="NEGARA2")
      colnames(tabel1_negara_migas)<-c("Negara",kolom_tahun_lalu, kolom_cumulative_tahun_lalu,kolom_bulan_lalu, kolom_bulan_ini, kolom_cumulative)
      tabel1_3 <- arrange(tabel1_negara_migas, desc(!!sym(kolom_cumulative)), desc(!!sym(kolom_cumulative_tahun_lalu)))
      
    } else {
      gabung1_negara_migas<-full_join(negara_migas_bulan_lalu,negara_migas_bulan_ini,by="NEGARA2")
      tabel1_3<-full_join(negara_migas_tahun_lalu,gabung1_negara_migas,by="NEGARA2")
      colnames(tabel1_3)<-c("Negara",kolom_tahun_lalu, kolom_bulan_lalu, kolom_bulan_ini)
      tabel1_3 <- arrange(tabel1_3, desc(!!sym(kolom_bulan_ini)), desc(!!sym(kolom_bulan_lalu)))
    }
    
    
    if(bulan!=1){
      tabel1_3<-mutate(tabel1_3, tabel1_3_mtm=100*(!!sym(kolom_bulan_ini)-!!sym(kolom_bulan_lalu))/!!sym(kolom_bulan_lalu), tabel1_3_yoy=100*(!!sym(kolom_bulan_ini)-!!sym(kolom_tahun_lalu))/!!sym(kolom_tahun_lalu), tabel1_3_ctc=100*(!!sym(kolom_cumulative)-!!sym(kolom_cumulative_tahun_lalu))/!!sym(kolom_cumulative_tahun_lalu), peran=100*!!sym(kolom_cumulative)/ekspor_cumulative_migas)
      colnames(tabel1_3)<-c("Negara",kolom_tahun_lalu, kolom_cumulative_tahun_lalu,kolom_bulan_lalu, kolom_bulan_ini, kolom_cumulative, "m-to-m (%)", "y-on-y (%)", "c-to-c (%)", paste("Peran (%)",kolom_cumulative))
    } else{
      tabel1_3<-mutate(tabel1_3, tabel1_3_mtm=100*(!!sym(kolom_bulan_ini)-!!sym(kolom_bulan_lalu))/!!sym(kolom_bulan_lalu), tabel1_3_yoy=100*(!!sym(kolom_bulan_ini)-!!sym(kolom_tahun_lalu))/!!sym(kolom_tahun_lalu),peran=100*!!sym(kolom_bulan_ini)/ekspor_bulan_ini_migas)
      colnames(tabel1_3)<-c("Negara",kolom_tahun_lalu,kolom_bulan_lalu, kolom_bulan_ini, "m-to-m (%)", "y-on-y (%)", paste("Peran (%)",kolom_bulan_ini))
    }
    
    tabel1_3 <- tabel1_3 %>%
      mutate(
        `m-to-m (%)` = ifelse(!is.na(!!sym(kolom_bulan_lalu)) & is.na(!!sym(kolom_bulan_ini)), -100, `m-to-m (%)`),
        `y-on-y (%)` = ifelse(!is.na(!!sym(kolom_tahun_lalu)) & is.na(!!sym(kolom_bulan_ini)), -100, `y-on-y (%)`)
      )
    
    if(bulan != 1){
      tabel1_3 <- tabel1_3 %>%
        mutate(
          `c-to-c (%)` = ifelse(!is.na(!!sym(kolom_cumulative_tahun_lalu)) & is.na(!!sym(kolom_cumulative)), -100, `c-to-c (%)`)
        )
    }
    
    tabel8<<-tabel1_3
    
    
    
    ##Tabel Output Nilai Ekspor Menurut Pelabuhan (Ringkasan Nilai/Volume Ekspor)
    pelabuhan_cumulative_tahun_lalu<-summarise(group_by(data[data$BULAN<=bulan & data$TAHUN==tahun-1,], PELABUHAN1), Total_Ekspor = sum(NILAI_JUTA))
    pelabuhan_cumulative<-summarise(group_by(data[data$BULAN<=bulan & data$TAHUN==tahun,], PELABUHAN1), Total_Ekspor = sum(NILAI_JUTA))
    pelabuhan_bulan_ini<-summarise(group_by(data[data$BULAN==bulan & data$TAHUN==tahun,], PELABUHAN1), Total_Ekspor = sum(NILAI_JUTA))
    pelabuhan_bulan_lalu<-summarise(group_by(data[data$BULAN==bulan_sebelumnya & data$TAHUN==tahun_sebelumnya,], PELABUHAN1), Total_Ekspor = sum(NILAI_JUTA))
    pelabuhan_tahun_lalu<-summarise(group_by(data[data$BULAN==bulan & data$TAHUN==tahun-1,], PELABUHAN1), Total_Ekspor = sum(NILAI_JUTA))
    
    
    if (bulan!=1) {
      gabung1_pelabuhan<-full_join(pelabuhan_bulan_ini,pelabuhan_cumulative,by="PELABUHAN1")
      gabung2_pelabuhan<-full_join(pelabuhan_bulan_lalu,gabung1_pelabuhan,by="PELABUHAN1")
      gabung3_pelabuhan<-full_join(pelabuhan_cumulative_tahun_lalu,gabung2_pelabuhan,by="PELABUHAN1")
      tabel1_pelabuhan<-full_join(pelabuhan_tahun_lalu,gabung3_pelabuhan,by="PELABUHAN1")
      colnames(tabel1_pelabuhan)<-c("Pelabuhan",kolom_tahun_lalu, kolom_cumulative_tahun_lalu,kolom_bulan_lalu, kolom_bulan_ini, kolom_cumulative)
      tabel1_4 <- arrange(tabel1_pelabuhan, desc(!!sym(kolom_cumulative)), desc(!!sym(kolom_cumulative_tahun_lalu)))
      
    } else {
      gabung1_pelabuhan<-full_join(pelabuhan_bulan_lalu,pelabuhan_bulan_ini,by="PELABUHAN1")
      tabel1_4<-full_join(pelabuhan_tahun_lalu,gabung1_pelabuhan,by="PELABUHAN1")
      colnames(tabel1_4)<-c("Pelabuhan",kolom_tahun_lalu, kolom_bulan_lalu, kolom_bulan_ini)
      tabel1_4 <- arrange(tabel1_4, desc(!!sym(kolom_bulan_ini)), desc(!!sym(kolom_bulan_lalu)))
    }
    
    
    if(bulan!=1){
      tabel1_4<-mutate(tabel1_4, tabel1_4_mtm=100*(!!sym(kolom_bulan_ini)-!!sym(kolom_bulan_lalu))/!!sym(kolom_bulan_lalu), tabel1_4_yoy=100*(!!sym(kolom_bulan_ini)-!!sym(kolom_tahun_lalu))/!!sym(kolom_tahun_lalu), tabel1_4_ctc=100*(!!sym(kolom_cumulative)-!!sym(kolom_cumulative_tahun_lalu))/!!sym(kolom_cumulative_tahun_lalu), peran=100*!!sym(kolom_cumulative)/ekspor_cumulative)
      colnames(tabel1_4)<-c("Pelabuhan",kolom_tahun_lalu, kolom_cumulative_tahun_lalu,kolom_bulan_lalu, kolom_bulan_ini, kolom_cumulative, "m-to-m (%)", "y-on-y (%)", "c-to-c (%)", paste("Peran (%)",kolom_cumulative))
    } else{
      tabel1_4<-mutate(tabel1_4, tabel1_4_mtm=100*(!!sym(kolom_bulan_ini)-!!sym(kolom_bulan_lalu))/!!sym(kolom_bulan_lalu), tabel1_4_yoy=100*(!!sym(kolom_bulan_ini)-!!sym(kolom_tahun_lalu))/!!sym(kolom_tahun_lalu),peran=100*!!sym(kolom_bulan_ini)/ekspor_bulan_ini)
      colnames(tabel1_4)<-c("Pelabuhan",kolom_tahun_lalu,kolom_bulan_lalu, kolom_bulan_ini, "m-to-m (%)", "y-on-y (%)", paste("Peran (%)",kolom_bulan_ini))
    }
    
    tabel1_4 <- tabel1_4 %>%
      mutate(
        `m-to-m (%)` = ifelse(!is.na(!!sym(kolom_bulan_lalu)) & is.na(!!sym(kolom_bulan_ini)), -100, `m-to-m (%)`),
        `y-on-y (%)` = ifelse(!is.na(!!sym(kolom_tahun_lalu)) & is.na(!!sym(kolom_bulan_ini)), -100, `y-on-y (%)`)
      )
    
    if(bulan != 1){
      tabel1_4 <- tabel1_4 %>%
        mutate(
          `c-to-c (%)` = ifelse(!is.na(!!sym(kolom_cumulative_tahun_lalu)) & is.na(!!sym(kolom_cumulative)), -100, `c-to-c (%)`)
        )
    }
    
    tabel2<<-tabel1_4
    
    
    
    ## Tabel Output Volume Ekspor Menurut Pelabuhan (Ringkasan Nilai/Volume Ekspor)
    
    # Total ekspor (cumulative dan bulan ini)
    ekspor_cumulative_volume <- sum(data$BERAT[data$BULAN <= bulan & data$TAHUN == tahun], na.rm = TRUE)/1000000
    ekspor_bulan_ini_volume <- sum(data$BERAT[data$BULAN == bulan & data$TAHUN == tahun], na.rm = TRUE)/1000000
    
    # Per pelabuhan
    pelabuhan_volume_cumulative_tahun_lalu <- data %>%
      filter(BULAN <= bulan, TAHUN == tahun - 1) %>%
      group_by(PELABUHAN1) %>%
      summarise(Total_Volume = sum(BERAT, na.rm = TRUE) / 1000000)
    
    pelabuhan_volume_cumulative <- data %>%
      filter(BULAN <= bulan, TAHUN == tahun) %>%
      group_by(PELABUHAN1) %>%
      summarise(Total_Volume = sum(BERAT, na.rm = TRUE) / 1000000)
    
    pelabuhan_volume_bulan_ini <- data %>%
      filter(BULAN == bulan, TAHUN == tahun) %>%
      group_by(PELABUHAN1) %>%
      summarise(Total_Volume = sum(BERAT, na.rm = TRUE) / 1000000)
    
    pelabuhan_volume_bulan_lalu <- data %>%
      filter(BULAN == bulan_sebelumnya, TAHUN == tahun_sebelumnya) %>%
      group_by(PELABUHAN1) %>%
      summarise(Total_Volume = sum(BERAT, na.rm = TRUE) / 1000000)
    
    pelabuhan_volume_tahun_lalu <- data %>%
      filter(BULAN == bulan, TAHUN == tahun - 1) %>%
      group_by(PELABUHAN1) %>%
      summarise(Total_Volume = sum(BERAT, na.rm = TRUE) / 1000000)
    
    # Penggabungan dan pengolahan
    if (bulan != 1) {
      gabung1_pelabuhan_volume <- full_join(pelabuhan_volume_bulan_ini, pelabuhan_volume_cumulative, by = "PELABUHAN1")
      gabung2_pelabuhan_volume <- full_join(pelabuhan_volume_bulan_lalu, gabung1_pelabuhan_volume, by = "PELABUHAN1")
      gabung3_pelabuhan_volume <- full_join(pelabuhan_volume_cumulative_tahun_lalu, gabung2_pelabuhan_volume, by = "PELABUHAN1")
      tabel1_pelabuhan_volume <- full_join(pelabuhan_volume_tahun_lalu, gabung3_pelabuhan_volume, by = "PELABUHAN1")
      
      # Penamaan kolom
      colnames(tabel1_pelabuhan_volume) <- c("Pelabuhan", kolom_tahun_lalu, kolom_cumulative_tahun_lalu,
                                             kolom_bulan_lalu, kolom_bulan_ini, kolom_cumulative)
      
      # Urutkan berdasarkan kumulatif
      tabel1_5 <- arrange(tabel1_pelabuhan_volume, desc(!!sym(kolom_cumulative)),desc(!!sym(kolom_cumulative_tahun_lalu)))
      
      # Tambahkan mtm, yoy, ctc, peran
      tabel1_5 <- tabel1_5 %>%
        mutate(
          `m-to-m (%)` = 100 * (!!sym(kolom_bulan_ini) - !!sym(kolom_bulan_lalu)) / !!sym(kolom_bulan_lalu),
          `y-on-y (%)` = 100 * (!!sym(kolom_bulan_ini) - !!sym(kolom_tahun_lalu)) / !!sym(kolom_tahun_lalu),
          `c-to-c (%)` = 100 * (!!sym(kolom_cumulative) - !!sym(kolom_cumulative_tahun_lalu)) / !!sym(kolom_cumulative_tahun_lalu),
          !!paste("Peran (%)", kolom_cumulative) := 100 * !!sym(kolom_cumulative) / ekspor_cumulative_volume
        )
      
    } else {
      gabung1_pelabuhan_volume <- full_join(pelabuhan_volume_bulan_lalu, pelabuhan_volume_bulan_ini, by = "PELABUHAN1")
      tabel1_5 <- full_join(pelabuhan_volume_tahun_lalu, gabung1_pelabuhan_volume, by = "PELABUHAN1")
      
      # Penamaan kolom
      colnames(tabel1_5) <- c("Pelabuhan", kolom_tahun_lalu, kolom_bulan_lalu, kolom_bulan_ini)
      
      # Urutkan berdasarkan bulan ini
      tabel1_5 <- arrange(tabel1_5, desc(!!sym(kolom_bulan_ini)), desc(!!sym(kolom_bulan_lalu)))
      
      # Tambahkan mtm, yoy, peran
      tabel1_5 <- tabel1_5 %>%
        mutate(
          `m-to-m (%)` = 100 * (!!sym(kolom_bulan_ini) - !!sym(kolom_bulan_lalu)) / !!sym(kolom_bulan_lalu),
          `y-on-y (%)` = 100 * (!!sym(kolom_bulan_ini) - !!sym(kolom_tahun_lalu)) / !!sym(kolom_tahun_lalu),
          !!paste("Peran (%)", kolom_bulan_ini) := 100 * !!sym(kolom_bulan_ini) / ekspor_bulan_ini_volume
        )
    }
    
    tabel1_5 <- tabel1_5 %>%
      mutate(
        `m-to-m (%)` = ifelse(!is.na(!!sym(kolom_bulan_lalu)) & is.na(!!sym(kolom_bulan_ini)), -100, `m-to-m (%)`),
        `y-on-y (%)` = ifelse(!is.na(!!sym(kolom_tahun_lalu)) & is.na(!!sym(kolom_bulan_ini)), -100, `y-on-y (%)`)
      )
    
    if(bulan != 1){
      tabel1_5 <- tabel1_5 %>%
        mutate(
          `c-to-c (%)` = ifelse(!is.na(!!sym(kolom_cumulative_tahun_lalu)) & is.na(!!sym(kolom_cumulative)), -100, `c-to-c (%)`)
        )
    }
    
    # Simpan hasil akhir
    tabel5 <<- tabel1_5
    
    
    
    ## ----------------------------------------------------------------- SLIDE 2 ---------------------------------------------------------------------------------------------
    rekap_HS2 <- summarise(group_by(data[data$BULAN==bulan & data$TAHUN==tahun & data$JENIS=="NON MIGAS",], HSDUA, HS2), Total_Ekspor = sum(NILAI_JUTA))
    rekap_HS2<-mutate(rekap_HS2,share=100*Total_Ekspor/ekspor_bulan_ini_nonmigas)
    rekap_HS2<-arrange(rekap_HS2,desc(share))
    tabel10_1<-rekap_HS2
    colnames(tabel10_1)<-c("Kode HS", "Deskripsi", paste("Nilai Ekspor ", kolom_bulan_ini), "Share (%)")
    tabel15<<-tabel10_1
    
    
    rekap_HS2_bulan_ini <- summarise(group_by(data[data$BULAN==bulan & data$TAHUN==tahun & data$JENIS=="NON MIGAS",], HSDUA, HS2), Total_Ekspor = sum(NILAI_JUTA))
    rekap_HS2_bulan_lalu <- summarise(group_by(data[data$BULAN==bulan_sebelumnya & data$TAHUN==tahun_sebelumnya & data$JENIS=="NON MIGAS",], HSDUA, HS2), Total_Ekspor = sum(NILAI_JUTA))
    
    
    rekap_HS2_gabung <- full_join(rekap_HS2_bulan_lalu, rekap_HS2_bulan_ini, by= "HSDUA")
    rekap_HS2_gabung <- rename(rekap_HS2_gabung, Ekspor_Bulan_Lalu=Total_Ekspor.x, Ekspor_Bulan_Ini=Total_Ekspor.y)
    
    rekap_HS2_gabung$HS2.x <- ifelse(is.na(rekap_HS2_gabung$HS2.x), rekap_HS2_gabung$HS2.y, rekap_HS2_gabung$HS2.x)
    rekap_HS2_gabung$HS2.y <- ifelse(is.na(rekap_HS2_gabung$HS2.y), rekap_HS2_gabung$HS2.x, rekap_HS2_gabung$HS2.y)
    
    rekap_HS2_gabung$Ekspor_Bulan_Lalu <- ifelse(is.na(rekap_HS2_gabung$Ekspor_Bulan_Lalu), 0, rekap_HS2_gabung$Ekspor_Bulan_Lalu)
    rekap_HS2_gabung$Ekspor_Bulan_Ini <- ifelse(is.na(rekap_HS2_gabung$Ekspor_Bulan_Ini), 0, rekap_HS2_gabung$Ekspor_Bulan_Ini)
    
    rekap_HS2_gabung <- select(rekap_HS2_gabung, -HS2.y)
    rekap_HS2_gabung <- rename(rekap_HS2_gabung, HS2=HS2.x)
    
    if ((bulan_sebelumnya %in% data$BULAN) & (tahun_sebelumnya %in% data$TAHUN)) {
    } else {
      rekap_HS2_gabung$Ekspor_Bulan_Lalu <- rep(NA_real_, nrow(rekap_HS2_gabung))
    }
    
    rekap_HS2_gabung <- mutate(rekap_HS2_gabung, peningkatan=Ekspor_Bulan_Ini-Ekspor_Bulan_Lalu)
    tabel9_1<<-arrange(rekap_HS2_gabung,desc(peningkatan))
    colnames(tabel9_1)<-c("Kode HS","Deskripsi",paste("Nilai Ekspor ", kolom_bulan_lalu),paste("Nilai Ekspor ", kolom_bulan_ini),"Peningkatan Nilai Ekspor")
    
    tabel14<<-tabel9_1
    
    ## ----------------------------------------------------------------- SLIDE 3 ---------------------------------------------------------------------------------------------
    
    rekap_negara <- summarise(group_by(data[data$BULAN==bulan & data$TAHUN==tahun,], NEGARA2), Total_Ekspor = sum(NILAI_JUTA))
    rekap_negara <- mutate(rekap_negara, share=100*Total_Ekspor/ekspor_bulan_ini)
    rekap_negara<-arrange(rekap_negara,desc(share))
    tabel4_1<-rekap_negara
    colnames(tabel4_1)<-c("Negara", paste("Nilai Ekspor ", kolom_bulan_ini), "Share (%)")
    tabel6<<-tabel4_1
    
    
    
    rekap_HS2_negara_utama <- summarise(group_by(data[data$BULAN==bulan & data$TAHUN==tahun & data$NEGARA2==rekap_negara$NEGARA2[1],], HSDUA, HS2), Total_Ekspor = sum(NILAI_JUTA))
    ekspor_bulan_ini_negara_utama <- rekap_negara$Total_Ekspor[1]
    rekap_HS2_negara_utama <- mutate(rekap_HS2_negara_utama, share=100*Total_Ekspor/ekspor_bulan_ini_negara_utama)
    rekap_HS2_negara_utama <- arrange(rekap_HS2_negara_utama,desc(share))
    tabel5_1<-rekap_HS2_negara_utama
    colnames(tabel5_1)<-c("Kode HS","Deskripsi", paste("Nilai Ekspor ", kolom_bulan_ini), "Share (%)")
    tabel9<<-tabel5_1
    
    
    ekspor_negara_utama_waktu <- summarise(group_by(data[data$NEGARA2==rekap_negara$NEGARA2[1],], BULAN, TAHUN), Total_Ekspor = sum(NILAI_JUTA))
    ekspor_negara_utama_waktu <- arrange(ekspor_negara_utama_waktu,TAHUN, BULAN)
    
    
    kumpulan_tahun <- sort(unique(data$TAHUN[data$TAHUN<=tahun]))
    
    tabel_ekspor_negara_utama_waktu<-matrix(nrow = 12,ncol = 1+length(kumpulan_tahun))
    colnames(tabel_ekspor_negara_utama_waktu) <- rep("", ncol(tabel_ekspor_negara_utama_waktu))
    
    for(i in 1:ncol(tabel_ekspor_negara_utama_waktu)){
      if(i==1){colnames(tabel_ekspor_negara_utama_waktu)[i] <- "Bulan"}
      else {colnames(tabel_ekspor_negara_utama_waktu)[i] <- paste("Nilai Ekspor ", kumpulan_tahun[i-1])}}
    
    tabel_ekspor_negara_utama_waktu<-as.data.frame(tabel_ekspor_negara_utama_waktu)
    tabel_ekspor_negara_utama_waktu$Bulan<-month.name
    
    for(i in 2:ncol(tabel_ekspor_negara_utama_waktu)){
      for(j in 1:12){
        tabel_ekspor_negara_utama_waktu[j,i]<- sum(ekspor_negara_utama_waktu$Total_Ekspor[ekspor_negara_utama_waktu$BULAN==j&ekspor_negara_utama_waktu$TAHUN==kumpulan_tahun[i-1]])
        if(j>bulan && i>=ncol(tabel_ekspor_negara_utama_waktu)){
          tabel_ekspor_negara_utama_waktu[j,i]<- NA}}}
    
    tabel10<<- tabel_ekspor_negara_utama_waktu
    
    rekap_negara_cumulative <- summarise(group_by(data[data$BULAN<=bulan & data$TAHUN==tahun,], NEGARA2), Total_Ekspor = sum(NILAI_JUTA))
    rekap_negara_cumulative <- mutate(rekap_negara_cumulative, share=100*Total_Ekspor/ekspor_cumulative)
    rekap_negara_cumulative<-arrange(rekap_negara_cumulative,desc(share))
    tabel7_1<-rekap_negara_cumulative
    colnames(tabel7_1)<-c("Negara", paste("Nilai Ekspor ", kolom_cumulative), "Share (%)")
    tabel11<<-tabel7_1
    
    tabel8_1 <- summarise(group_by(data[data$BULAN <= bulan & data$NEGARA2 == rekap_negara$NEGARA2[1] & data$TAHUN<=tahun, ],TAHUN),Total_Ekspor = sum(NILAI_JUTA, na.rm = TRUE))
    tabel8_1<-arrange(tabel8_1, TAHUN)
    tabel8_1$TAHUN<-as.character(tabel8_1$TAHUN)
    colnames(tabel8_1)<-c(paste("Periode ",month.abb[1]," - ",month.abb[bulan]), "Nilai Ekspor Kumulatif")
    tabel12<<-tabel8_1
    
    ## ----------------------------------------------------------------- SLIDE 4 ---------------------------------------------------------------------------------------------
    perkembangan_ekspor <- summarise(group_by(data, BULAN, TAHUN), Total_Ekspor = sum(NILAI_JUTA))
    perkembangan_ekspor <- arrange(perkembangan_ekspor,TAHUN, BULAN)
    
    
    tabel_perkembangan_ekspor<-matrix(nrow = 12,ncol = 1+length(kumpulan_tahun))
    colnames(tabel_perkembangan_ekspor) <- rep("", ncol(tabel_perkembangan_ekspor))
    
    for(i in 1:ncol(tabel_perkembangan_ekspor)){
      if(i==1){colnames(tabel_perkembangan_ekspor)[i] <- "Bulan"}
      else {colnames(tabel_perkembangan_ekspor)[i] <- paste("Nilai Ekspor ",kumpulan_tahun[i-1])}}
    
    tabel_perkembangan_ekspor<-as.data.frame(tabel_perkembangan_ekspor)
    tabel_perkembangan_ekspor$Bulan<-month.name
    
    for(i in 2:ncol(tabel_perkembangan_ekspor)){
      for(j in 1:12){
        tabel_perkembangan_ekspor[j,i]<- sum(perkembangan_ekspor$Total_Ekspor[perkembangan_ekspor$BULAN==j&perkembangan_ekspor$TAHUN==kumpulan_tahun[i-1]])
        if(j>bulan && i>=ncol(tabel_perkembangan_ekspor)){
          tabel_perkembangan_ekspor[j,i]<- NA}}}
    
    tabel3<<-tabel_perkembangan_ekspor
    
    
    
    tabel3_1 <- summarise(group_by(data[data$BULAN <= bulan  & data$TAHUN<=tahun, ],TAHUN),Total_Ekspor = sum(NILAI_JUTA, na.rm = TRUE))
    tabel3_1<-arrange(tabel3_1, TAHUN)
    tabel3_1$TAHUN<-as.character(tabel3_1$TAHUN)
    colnames(tabel3_1)<-c(paste("Periode ",month.abb[1]," - ",month.abb[bulan]), "Nilai Ekspor Kumulatif")
    tabel4<<-tabel3_1
    
    
    # Reaktif data berdasarkan pilihan
    datasetInput <- reactive({
      switch(input$pilih_tabel,
             "1. Nilai Ekspor Menurut Sektor"=tabel1,
             "2. Nilai Ekspor Menurut Pelabuhan"=tabel2,
             "3. Perkembangan Nilai Ekspor"=tabel3,
             "4. Perkembangan Nilai Ekspor (c-t-c)"=tabel4,
             "5. Volume Ekspor Menurut Pelabuhan"=tabel5,
             "6. Nilai Ekspor Menurut Negara Tujuan"=tabel6,
             "7. Nilai Ekspor Nonmigas Menurut Negara Tujuan"=tabel7,
             "8. Nilai Ekspor Migas Menurut Negara Tujuan"=tabel8,
             "9. Nilai Ekspor Negara Tujuan Utama HS2 Digit"=tabel9,
             "10. Perkembangan Nilai Ekspor Negara Tujuan Utama"=tabel10,
             "11. Nilai Ekspor Kumulatif Menurut Negara Tujuan"=tabel11,
             "12. Perkembangan Nilai Ekspor Negara Tujuan Utama (c-t-c)"=tabel12,
             "13. Nilai Ekspor Nonmigas Menurut Golongan Barang HS2 Digit"=tabel13,
             "14. Peningkatan/Penurunan Nilai Ekspor Nonmigas HS2 Digit (m-t-m)"=tabel14,
             "15. Share Nilai Ekspor Nonmigas HS2 Digit"=tabel15
      )
    })
    
    
    output$downloadData <- downloadHandler(
      filename = function() {
        paste0(input$pilih_tabel, ".xlsx")
      },
      content = function(file) {
        data_to_download <- datasetInput()
        
        if (is.null(data_to_download) || nrow(data_to_download) == 0) {
          data_to_download <- data.frame(Pesan = "Data tidak tersedia", check.names = FALSE)
        } else {
          # ======== PERBAIKAN NAMA KOLOM ========
          clean_colnames <- function(names_vec) {
            names_vec <- gsub("\\s+", " ", names_vec)         # Hapus spasi berlebih
            names_vec <- trimws(names_vec)                    # Trim kiri-kanan
            # Hanya ubah simbol yang berlebihan
            names_vec <- gsub("(?<![a-zA-Z0-9])-\\s*-\\s*(?![a-zA-Z0-9])", " - ", names_vec, perl = TRUE)  # Perbaiki - yang berlebihan
            return(names_vec)
          }
          
          colnames(data_to_download) <- clean_colnames(colnames(data_to_download))
          
          # ======== FORMAT ANGKA 2 DIGIT ========
          numeric_cols <- sapply(data_to_download, is.numeric)
          
          data_to_download[numeric_cols] <- lapply(
            data_to_download[numeric_cols],
            function(x) {
              out <- ifelse(
                is.na(x), "NA",
                ifelse(x == 0, "ZERO", formatC(x, format = "f", digits = 2, big.mark = ".", decimal.mark = ","))
              )
              return(out)
            }
          )
          
          
          # Ganti "NA" dan "ZERO" dengan "-"
          data_to_download <- data.frame(
            lapply(data_to_download, function(col) {
              col <- as.character(col)
              col[col %in% c("NA", "ZERO")] <- "-"
              return(col)
            }),
            stringsAsFactors = FALSE,
            check.names = FALSE
          )
          
        }
        
        # ======== MEMBUAT FILE EXCEL ========
        wb <- openxlsx::createWorkbook()
        openxlsx::addWorksheet(wb, "Sheet1")
        
        # Style judul
        title_style <- openxlsx::createStyle(
          fontSize = 14,
          textDecoration = "bold",
          halign = "center"
        )
        
        # Style header (center horizontal & vertical, wrap, border)
        header_style <- openxlsx::createStyle(
          textDecoration = "bold",
          wrapText = TRUE,
          halign = "center",
          valign = "center",
          border = "TopBottomLeftRight",
          borderStyle = "thin"
        )
        
        # Style isi data
        data_style <- openxlsx::createStyle(
          border = "TopBottomLeftRight",
          borderStyle = "thin"
        )
        
        # Tulis judul tabel
        # Buat teks judul
        table_title <- paste("Tabel", input$pilih_tabel)
        
        # Style tanpa wrap dan tanpa merge
        title_style <- openxlsx::createStyle(
          textDecoration = "bold",
          halign = "left",
          valign = "center",
          fontSize = 12,
          wrapText = FALSE
        )
        
        # Tulis judul di sel A1 saja
        openxlsx::writeData(wb, sheet = 1, x = table_title, startRow = 1, startCol = 1, colNames = FALSE)
        openxlsx::addStyle(wb, sheet = 1, style = title_style, rows = 1, cols = 1)
        
        
        # Tulis data (mulai baris ke-2)
        openxlsx::writeData(wb, sheet = 1, x = data_to_download, startRow = 2, headerStyle = header_style)
        
        # Tambahkan border ke isi data
        openxlsx::addStyle(
          wb, sheet = 1, style = data_style,
          rows = 3:(nrow(data_to_download) + 2),
          cols = 1:ncol(data_to_download),
          gridExpand = TRUE,
          stack = TRUE
        )
        
        # ======== TAMBAHKAN CATATAN DI BAWAH TABEL ========
        # Ambil waktu saat ini
        timestamp <- format(Sys.time(), "%d %B %Y pukul %H:%M WIB", tz = "Asia/Jakarta")
        
        # Teks catatan
        footer_text <- paste("Sumber: BPS Kabupaten Karimun (data diolah), diakses pada", timestamp)
        
        # Tentukan baris tempat catatan ditulis (baris setelah data terakhir + 2)
        footer_row <- nrow(data_to_download) + 4  # 1 baris kosong setelah tabel
        
        # Tulis catatan di kolom pertama
        openxlsx::writeData(wb, sheet = 1, x = footer_text, startRow = footer_row, startCol = 1, colNames = FALSE)
        
        # Tambahkan style miring & kecil
        footer_style <- openxlsx::createStyle(
          fontSize = 9,
          textDecoration = "italic",
          halign = "left"
        )
        
        openxlsx::addStyle(wb, sheet = 1, style = footer_style, rows = footer_row, cols = 1, gridExpand = TRUE)
        
        
        # Atur lebar kolom
        get_width <- function(col) {
          max_len <- max(nchar(as.character(col)), na.rm = TRUE)
          return(min(max(10, max_len + 2), 30))
        }
        
        col_widths <- sapply(data_to_download, get_width)
        openxlsx::setColWidths(wb, sheet = 1, cols = 1:ncol(data_to_download), widths = col_widths)
        
        # Simpan
        openxlsx::saveWorkbook(wb, file, overwrite = TRUE)
      }
    )
    
    
    ## Download semua tabel
    output$downloadAll <- downloadHandler(
      filename = function() {
        "Semua Tabel Ekspor.zip"
      },
      content = function(file) {
        # Buat folder sementara
        temp_dir <- tempdir()
        
        title_style <- openxlsx::createStyle(
          fontSize = 14,
          textDecoration = "bold",
          halign = "center"
        )
        
        # Daftar nama file dan tabel beserta judul tabel
        daftar_tabel <- list(
          "1. Nilai Ekspor Menurut Sektor.xlsx" = list(tabel = tabel1, judul = "Tabel 1. Nilai Ekspor Menurut Sektor"),
          "2. Nilai Ekspor Menurut Pelabuhan.xlsx" = list(tabel = tabel2, judul = "Tabel 2. Nilai Ekspor Menurut Pelabuhan"),
          "3. Perkembangan Nilai Ekspor.xlsx" = list(tabel = tabel3, judul = "Tabel 3. Perkembangan Nilai Ekspor"),
          "4. Perkembangan Nilai Ekspor (c-t-c).xlsx" = list(tabel = tabel4, judul = "Tabel 4. Perkembangan Nilai Ekspor (c-t-c)"),
          "5. Volume Ekspor Menurut Pelabuhan.xlsx" = list(tabel = tabel5, judul = "Tabel 5. Volume Ekspor Menurut Pelabuhan"),
          "6. Nilai Ekspor Menurut Negara Tujuan.xlsx" = list(tabel = tabel6, judul = "Tabel 6. Nilai Ekspor Menurut Negara Tujuan"),
          "7. Nilai Ekspor Nonmigas Menurut Negara Tujuan.xlsx" = list(tabel = tabel7, judul = "Tabel 7. Nilai Ekspor Nonmigas Menurut Negara Tujuan"),
          "8. Nilai Ekspor Migas Menurut Negara Tujuan.xlsx" = list(tabel = tabel8, judul = "Tabel 8. Nilai Ekspor Migas Menurut Negara Tujuan"),
          "9. Nilai Ekspor Negara Tujuan Utama HS2 Digit.xlsx" = list(tabel = tabel9, judul = "Tabel 9. Nilai Ekspor Negara Tujuan Utama HS2 Digit"),
          "10. Perkembangan Nilai Ekspor Negara Tujuan Utama.xlsx" = list(tabel = tabel10, judul = "Tabel 10. Perkembangan Nilai Ekspor Negara Tujuan Utama"),
          "11. Nilai Ekspor Kumulatif Menurut Negara Tujuan.xlsx" = list(tabel = tabel11, judul = "Tabel 11. Nilai Ekspor Kumulatif Menurut Negara Tujuan"),
          "12. Perkembangan Nilai Ekspor Negara Tujuan Utama (c-t-c).xlsx" = list(tabel = tabel12, judul = "Tabel 12. Perkembangan Nilai Ekspor Negara Tujuan Utama (c-t-c)"),
          "13. Nilai Ekspor Nonmigas Menurut Golongan Barang HS2 Digit.xlsx" = list(tabel = tabel13, judul = "Tabel 13. Nilai Ekspor Nonmigas Menurut Golongan Barang HS2 Digit"),
          "14. Peningkatan_Penurunan Nilai Ekspor Nonmigas HS2 Digit (m-t-m).xlsx" = list(tabel = tabel14, judul = "Tabel 14. Peningkatan/Penurunan Nilai Ekspor Nonmigas HS2 Digit (m-t-m)"),
          "15. Share Nilai Ekspor Nonmigas HS2 Digit.xlsx" = list(tabel = tabel15, judul = "Tabel 15. Share Nilai Ekspor Nonmigas HS2 Digit")
        )
        
        clean_colnames <- function(names_vec) {
          names_vec <- gsub("\\s+", " ", names_vec)
          names_vec <- trimws(names_vec)
          names_vec <- gsub("(?<![a-zA-Z0-9])-\\s*-\\s*(?![a-zA-Z0-9])", " - ", names_vec, perl = TRUE)
          return(names_vec)
        }
        
        format_numeric <- function(df) {
          numeric_cols <- sapply(df, is.numeric)
          
          df[numeric_cols] <- lapply(df[numeric_cols], function(x) {
            out <- ifelse(
              is.na(x), "NA",
              ifelse(x == 0, "ZERO", formatC(x, format = "f", digits = 2, big.mark = ".", decimal.mark = ","))
            )
            return(out)
          })
          
          
          # Ganti "NA" dan "ZERO" dengan "-"
          df <- data.frame(
            lapply(df, function(col) {
              col <- as.character(col)
              col[col %in% c("NA", "ZERO")] <- "-"
              return(col)
            }),
            stringsAsFactors = FALSE,
            check.names = FALSE
          )
          
          return(df)
        }
        
        
        write_table_to_excel <- function(data, file_path, judul_tabel) {
          colnames(data) <- clean_colnames(colnames(data))
          data <- format_numeric(data)
          
          wb <- openxlsx::createWorkbook()
          openxlsx::addWorksheet(wb, "Sheet1")
          
          header_style <- openxlsx::createStyle(
            textDecoration = "bold",
            wrapText = TRUE,
            halign = "center",
            valign = "center",
            border = "TopBottomLeftRight",
            borderStyle = "thin"
          )
          
          data_style <- openxlsx::createStyle(
            border = "TopBottomLeftRight",
            borderStyle = "thin"
          )
          
          judul_style <- openxlsx::createStyle(
            textDecoration = "bold",
            halign = "left",
            valign = "center",
            fontSize = 12,
            wrapText = FALSE
          )
          
          openxlsx::writeData(wb, sheet = 1, x = judul_tabel, startRow = 1, startCol = 1, colNames = FALSE)
          openxlsx::addStyle(wb, sheet = 1, style = judul_style, rows = 1, cols = 1)
          
          openxlsx::writeData(wb, sheet = 1, x = data, startRow = 2, headerStyle = header_style)
          openxlsx::addStyle(
            wb, sheet = 1, style = data_style,
            rows = 3:(nrow(data) + 2),
            cols = 1:ncol(data),
            gridExpand = TRUE,
            stack = TRUE
          )
          
          # ===== Tambahkan footer =====
          timestamp <- format(Sys.time(), "%d %B %Y pukul %H:%M WIB", tz = "Asia/Jakarta")
          footer_text <- paste("Sumber: BPS Kabupaten Karimun (data diolah), diakses pada", timestamp)
          footer_row <- nrow(data) + 4
          
          openxlsx::writeData(wb, sheet = 1, x = footer_text, startRow = footer_row, startCol = 1, colNames = FALSE)
          footer_style <- openxlsx::createStyle(
            fontSize = 9,
            textDecoration = "italic",
            halign = "left"
          )
          openxlsx::addStyle(wb, sheet = 1, style = footer_style, rows = footer_row, cols = 1)
          
          get_width <- function(col) {
            max_len <- max(nchar(as.character(col)), na.rm = TRUE)
            return(min(max(10, max_len + 2), 30))
          }
          col_widths <- sapply(data, get_width)
          openxlsx::setColWidths(wb, sheet = 1, cols = 1:ncol(data), widths = col_widths)
          
          openxlsx::saveWorkbook(wb, file_path, overwrite = TRUE)
        }
        
        file_paths <- c()
        for (nama_file in names(daftar_tabel)) {
          path_file <- file.path(temp_dir, nama_file)
          write_table_to_excel(daftar_tabel[[nama_file]]$tabel, path_file, daftar_tabel[[nama_file]]$judul)
          file_paths <- c(file_paths, path_file)
        }
        
        zip::zipr(zipfile = file, files = file_paths)
      },
      contentType = "application/zip"  # <--- INI PENTING!
    )
    
    
    
    output$tabel_data_ekspor <- renderDT({
      req(data)  # Pastikan data tersedia
      
      data_sorted <- data %>%
        arrange(desc(TAHUN), desc(BULAN)) %>%
        select(-NILAI_JUTA)
      
      datatable(
        data_sorted,
        options = list(
          scrollX = TRUE,
          scrollY = "300px",
          fixedHeader = TRUE,
          ordering = TRUE,
          columnDefs = list(
            list(targets = 0, className = "dt-nowrap")
          )
        ),
        rownames = FALSE
      )
    })
    
    
    
    output$tabel_output_ekspor_3 <- renderDT({
      
      data_terpilih <- switch(input$pilih_tabel,
                              "1. Nilai Ekspor Menurut Sektor"=tabel1,
                              "2. Nilai Ekspor Menurut Pelabuhan"=tabel2,
                              "3. Perkembangan Nilai Ekspor"=tabel3,
                              "4. Perkembangan Nilai Ekspor (c-t-c)"=tabel4,
                              "5. Volume Ekspor Menurut Pelabuhan"=tabel5,
                              "6. Nilai Ekspor Menurut Negara Tujuan"=tabel6,
                              "7. Nilai Ekspor Nonmigas Menurut Negara Tujuan"=tabel7,
                              "8. Nilai Ekspor Migas Menurut Negara Tujuan"=tabel8,
                              "9. Nilai Ekspor Negara Tujuan Utama HS2 Digit"=tabel9,
                              "10. Perkembangan Nilai Ekspor Negara Tujuan Utama"=tabel10,
                              "11. Nilai Ekspor Kumulatif Menurut Negara Tujuan"=tabel11,
                              "12. Perkembangan Nilai Ekspor Negara Tujuan Utama (c-t-c)"=tabel12,
                              "13. Nilai Ekspor Nonmigas Menurut Golongan Barang HS2 Digit"=tabel13,
                              "14. Peningkatan/Penurunan Nilai Ekspor Nonmigas HS2 Digit (m-t-m)"=tabel14,
                              "15. Share Nilai Ekspor Nonmigas HS2 Digit"=tabel15
      )
      
      # 1. Format numerik, ubah NA jadi "NA", dan tandai 0 sebagai "ZERO"
      numerik_cols <- sapply(data_terpilih, is.numeric)
      data_terpilih[numerik_cols] <- lapply(
        data_terpilih[numerik_cols],
        function(x) {
          out <- ifelse(is.na(x), "NA",
                        ifelse(x == 0, "ZERO", formatC(x, format = "f", digits = 2, big.mark = ".", decimal.mark = ",")))
          return(out)
        }
      )
      
      # 2. Ganti "NA" dan "ZERO" menjadi "-"
      data_terpilih <- data.frame(
        lapply(data_terpilih, function(col) {
          col <- as.character(col)
          col[col %in% c("NA", "ZERO")] <- "-"
          return(col)
        }),
        stringsAsFactors = FALSE,
        check.names = FALSE
      )
      
      
      
      
      
      
      datatable(
        data_terpilih,
        options = list(
          scrollX = TRUE,
          scrollY = "300px",
          scrollCollapse = TRUE,  # Tambahkan ini agar tinggi menyesuaikan jumlah baris
          fixedHeader = TRUE,
          ordering = FALSE,
          columnDefs = list(
            list(targets = 0, className = "dt-nowrap")  # Hindari wrap teks kolom pertama
          )
        ),
        rownames = FALSE
      )
      
      
    })
    
  })
  
  
  
  ##-------------------------------------------Impor----------------------------------------------------------
  
  data_reaktif_impor <- reactiveVal(NULL)
  
  # Saat file di-upload, baca datanya dan update pilihan tahun
  observeEvent(input$file_input_impor, {
    req(input$file_input_impor)
    
    df <- tryCatch(
      as.data.frame(read_sav(input$file_input_impor$datapath)),
      error = function(e) {
        showNotification("Gagal membaca file. Pastikan format file adalah .sav", type = "error")
        return(NULL)
      }
    )
    
    req(df)
    
    # Cek apakah kolom TAHUN dan BULAN ada
    if (any(c("TAHUN", "BULAN", "NILAI", "BERAT", "PELABUHAN", 
              "HS", "NEGARA", "JENIS", "SEKTOR", "KOMODITI", 
              "HS2", "HSDUA", "NEGARA2", "PELABUHAN1") %in% names(df) == FALSE)) {
      showNotification("Ada kolom penting yang tidak ditemukan di data impor", type = "error")
      return()
    }
    
    
    data_reaktif_impor(df)  # Simpan ke reactive
    
    # Update pilihan tahun
    tahun_unik <- sort(unique(df$TAHUN))
    updateSelectInput(session, "tahun_impor", choices = tahun_unik)
    
    # Kosongkan pilihan bulan
    updateSelectInput(session, "bulan_impor", choices = NULL)
  })
  
  # Saat tahun dipilih, update bulan
  observeEvent(input$tahun_impor, {
    req(data_reaktif_impor())
    req(input$tahun_impor)
    
    df <- data_reaktif_impor()
    
    bulan_angka <- df %>%
      filter(TAHUN == input$tahun_impor) %>%
      pull(BULAN) %>%
      as.numeric() %>%
      unique() %>%
      sort()
    
    bulan_angka <- bulan_angka[!is.na(bulan_angka) & bulan_angka %in% 1:12]
    nama_bulan <- month.name[bulan_angka]
    names(bulan_angka) <- nama_bulan
    
    updateSelectInput(session, "bulan_impor", choices = bulan_angka)
  })
  
  
  
  observeEvent(input$analisis_button_impor, {
    
    req(data_reaktif_impor())
    req(input$tahun_impor)
    req(input$bulan_impor)
    
    bulan<-as.numeric(input$bulan_impor)
    tahun <- as.numeric(input$tahun_impor) 
    data <- data_reaktif_impor()
    
    
    data$BULAN<-as.numeric(data$BULAN)
    data$TAHUN<-as.numeric(data$TAHUN)
    
    data$PROPINSI<-as.character(data$PROPINSI)
    data$PROPINSI <- trimws(data$PROPINSI, which = "both")
    
    data$PELABUHAN<-as.character(data$PELABUHAN)
    data$PELABUHAN <- trimws(data$PELABUHAN, which = "both")
    
    data$HS<-as.character(data$HS)
    data$HS<-ifelse(nchar(data$HS)==7, paste0("0",data$HS), data$HS)
    data$HS <- trimws(data$HS, which = "both")
    
    data$NEGARA<-as.character(data$NEGARA)
    data$NEGARA <- trimws(data$NEGARA, which = "both")
    
    data$BERAT<-as.numeric(data$BERAT)
    data$NILAI<-as.numeric(data$NILAI)
    
    data$HSDUA<-as.character(data$HSDUA)
    data$HSDUA<-ifelse(nchar(data$HSDUA)==1, paste0("0",data$HSDUA), data$HSDUA)
    data$HSDUA <- trimws(data$HSDUA, which = "both")
    
    data$KABKOT<-as.character(data$KABKOT)
    data$KABKOT <- trimws(data$KABKOT, which = "both")
    
    data$PELABUHAN1<-as.character(data$PELABUHAN1)
    data$PELABUHAN1 <- trimws(data$PELABUHAN1, which = "both")
    
    data$NEGARA2<-as.character(data$NEGARA2)
    data$NEGARA2 <- trimws(data$NEGARA2, which = "both")
    
    data$HS2<-as.character(data$HS2)
    data$HS2 <- trimws(data$HS2, which = "both")
    
    data$JENIS<-as.character(data$JENIS)
    data$JENIS <- trimws(data$JENIS, which = "both")
    
    data$SEKTOR<-as.character(data$SEKTOR)
    data$SEKTOR <- trimws(data$SEKTOR, which = "both")
    
    data$KOMODITI<-as.character(data$KOMODITI)
    data$KOMODITI <- trimws(data$KOMODITI, which = "both")
    
    data<-mutate(data,NILAI_JUTA=NILAI/1000000)
    ##data <- data[data$PROVASAL == "21", ]
    data$NILAI_JUTA<-as.numeric(data$NILAI_JUTA)
    
    
    ## ----------------------------------------------------------------- SLIDE 1 ---------------------------------------------------------------------------------------------
    ##Ekspor Bulan Ini
    ekspor_bulan_ini<-sum(data$NILAI_JUTA[data$BULAN==bulan & data$TAHUN==tahun])
    ekspor_bulan_ini_migas<-sum(data$NILAI_JUTA[data$JENIS=="MIGAS" & data$BULAN==bulan & data$TAHUN==tahun])
    ekspor_bulan_ini_nonmigas<-sum(data$NILAI_JUTA[data$JENIS=="NON MIGAS" & data$BULAN==bulan & data$TAHUN==tahun])
    kolom_bulan_ini<-paste(month.abb[bulan]," ", tahun)
    
    ##Ekspor Bulan Sebelumnya (M to M)
    bulan_sebelumnya <- ifelse(bulan-1 == 0, 12, bulan-1)
    tahun_sebelumnya <- ifelse(bulan_sebelumnya == 12, tahun-1, tahun)
    
    ekspor_bulan_lalu<-sum(data$NILAI_JUTA[data$BULAN==bulan_sebelumnya & data$TAHUN==tahun_sebelumnya])
    ekspor_bulan_lalu_migas<-sum(data$NILAI_JUTA[data$JENIS=="MIGAS" & data$BULAN==bulan_sebelumnya & data$TAHUN==tahun_sebelumnya])
    ekspor_bulan_lalu_nonmigas<-sum(data$NILAI_JUTA[data$JENIS=="NON MIGAS" & data$BULAN==bulan_sebelumnya & data$TAHUN==tahun_sebelumnya])
    
    kolom_bulan_lalu<-paste(month.abb[bulan_sebelumnya]," ", tahun_sebelumnya)
    
    ##Ekspor Tahun Sebelumnya (Y on Y)
    ekspor_tahun_lalu<-sum(data$NILAI_JUTA[data$BULAN==bulan & data$TAHUN==tahun-1])
    ekspor_tahun_lalu_migas<-sum(data$NILAI_JUTA[data$JENIS=="MIGAS" & data$BULAN==bulan & data$TAHUN==tahun-1])
    ekspor_tahun_lalu_nonmigas<-sum(data$NILAI_JUTA[data$JENIS=="NON MIGAS" & data$BULAN==bulan & data$TAHUN==tahun-1])
    
    kolom_tahun_lalu<-paste(month.abb[bulan]," ", tahun-1)
    
    
    ##Ekspor Cumulative (C to C)
    ekspor_cumulative<-sum(data$NILAI_JUTA[data$BULAN<=bulan & data$TAHUN==tahun])
    ekspor_cumulative_migas<-sum(data$NILAI_JUTA[data$JENIS=="MIGAS" & data$BULAN<=bulan & data$TAHUN==tahun])
    ekspor_cumulative_nonmigas<-sum(data$NILAI_JUTA[data$JENIS=="NON MIGAS" & data$BULAN<=bulan & data$TAHUN==tahun])
    
    ekspor_cumulative_tahun_lalu<-sum(data$NILAI_JUTA[data$BULAN<=bulan & data$TAHUN==tahun-1])
    ekspor_cumulative_migas_tahun_lalu<-sum(data$NILAI_JUTA[data$JENIS=="MIGAS" & data$BULAN<=bulan & data$TAHUN==tahun-1])
    ekspor_cumulative_nonmigas_tahun_lalu<-sum(data$NILAI_JUTA[data$JENIS=="NON MIGAS" & data$BULAN<=bulan & data$TAHUN==tahun-1])
    
    kolom_cumulative<-paste(month.abb[1],"-", month.abb[bulan]," ", tahun)
    kolom_cumulative_tahun_lalu<-paste(month.abb[1],"-", month.abb[bulan]," ", tahun-1)
    
    
    ##Tabel Output
    sektor_cumulative_tahun_lalu<-summarise(group_by(data[data$BULAN<=bulan & data$TAHUN==tahun-1,], JENIS,SEKTOR), Total_Ekspor = sum(NILAI_JUTA))
    sektor_cumulative<-summarise(group_by(data[data$BULAN<=bulan & data$TAHUN==tahun,], JENIS,SEKTOR), Total_Ekspor = sum(NILAI_JUTA))
    sektor_bulan_ini<-summarise(group_by(data[data$BULAN==bulan & data$TAHUN==tahun,], JENIS,SEKTOR), Total_Ekspor = sum(NILAI_JUTA))
    sektor_bulan_lalu<-summarise(group_by(data[data$BULAN==bulan_sebelumnya & data$TAHUN==tahun_sebelumnya,], JENIS,SEKTOR), Total_Ekspor = sum(NILAI_JUTA))
    sektor_tahun_lalu<-summarise(group_by(data[data$BULAN==bulan & data$TAHUN==tahun-1,], JENIS,SEKTOR), Total_Ekspor = sum(NILAI_JUTA))
    
    
    if (bulan!=1) {
      gabung1<-full_join(sektor_bulan_ini,sektor_cumulative,by=c("JENIS","SEKTOR"))
      gabung2<-full_join(sektor_bulan_lalu,gabung1,by=c("JENIS","SEKTOR"))
      gabung3<-full_join(sektor_cumulative_tahun_lalu,gabung2,by=c("JENIS","SEKTOR"))
      tabel1<-full_join(sektor_tahun_lalu,gabung3,by=c("JENIS","SEKTOR"))
      colnames(tabel1)<-c("Jenis","Uraian",kolom_tahun_lalu, kolom_cumulative_tahun_lalu,kolom_bulan_lalu, kolom_bulan_ini, kolom_cumulative)
      tabel1 <- arrange(tabel1, desc(!!sym(kolom_cumulative)),desc(!!sym(kolom_cumulative_tahun_lalu)))
      
      tabel1_gabung<-data.frame(Jenis=c("","",""),Uraian=c("TOTAL IMPOR", "MIGAS", "NON MIGAS"),kolom_tahun_lalu=c(ekspor_tahun_lalu,ekspor_tahun_lalu_migas,ekspor_tahun_lalu_nonmigas), kolom_cumulative_tahun_lalu=c(ekspor_cumulative_tahun_lalu,ekspor_cumulative_migas_tahun_lalu,ekspor_cumulative_nonmigas_tahun_lalu),kolom_bulan_lalu=c(ekspor_bulan_lalu,ekspor_bulan_lalu_migas,ekspor_bulan_lalu_nonmigas),kolom_bulan_ini=c(ekspor_bulan_ini,ekspor_bulan_ini_migas,ekspor_bulan_ini_nonmigas), kolom_cumulative=c(ekspor_cumulative,ekspor_cumulative_migas,ekspor_cumulative_nonmigas)) 
      names(tabel1_gabung) <- names(tabel1)
      tabel1<-rbind(tabel1_gabung,tabel1)
      
    } else {
      gabung1<-full_join(sektor_bulan_lalu,sektor_bulan_ini,by=c("JENIS","SEKTOR"))
      tabel1<-full_join(sektor_tahun_lalu,gabung1,by=c("JENIS","SEKTOR"))
      colnames(tabel1)<-c("Jenis","Uraian",kolom_tahun_lalu, kolom_bulan_lalu, kolom_bulan_ini)
      tabel1 <- arrange(tabel1, desc(!!sym(kolom_bulan_ini)),desc(!!sym(kolom_bulan_lalu)))
      
      tabel1_gabung<-data.frame(Jenis=c("","",""),Uraian=c("TOTAL IMPOR", "MIGAS", "NON MIGAS"),kolom_tahun_lalu=c(ekspor_tahun_lalu,ekspor_tahun_lalu_migas,ekspor_tahun_lalu_nonmigas), kolom_bulan_lalu=c(ekspor_bulan_lalu,ekspor_bulan_lalu_migas,ekspor_bulan_lalu_nonmigas),kolom_bulan_ini=c(ekspor_bulan_ini,ekspor_bulan_ini_migas,ekspor_bulan_ini_nonmigas))      
      names(tabel1_gabung) <- names(tabel1)
      tabel1<-rbind(tabel1_gabung,tabel1)
    }
    
    tabel1$urutan <- NA
    
    # Urutan utama
    tabel1$urutan[tabel1$Uraian == "TOTAL IMPOR"] <- 1
    tabel1$urutan[tabel1$Uraian == "MIGAS"] <- 2
    
    # Urut sektor MIGAS
    i <- 3
    for (j in 1:nrow(tabel1)) {
      if (tabel1$Jenis[j] == "MIGAS" & !(tabel1$Uraian[j] %in% c("TOTAL IMPOR", "MIGAS", "NON MIGAS"))) {
        tabel1$urutan[j] <- i
        i <- i + 1
      }
    }
    
    # Urutan NON MIGAS
    tabel1$urutan[tabel1$Uraian == "NON MIGAS"] <- i
    i <- i + 1
    
    # Urut sektor NON MIGAS
    for (j in 1:nrow(tabel1)) {
      if (tabel1$Jenis[j] == "NON MIGAS" & !(tabel1$Uraian[j] %in% c("TOTAL IMPOR", "MIGAS", "NON MIGAS"))) {
        tabel1$urutan[j] <- i
        i <- i + 1
      }
    }
    
    # Urutkan
    tabel1 <- tabel1[order(tabel1$urutan), ]
    tabel1$Uraian <- ifelse(!(tabel1$Uraian %in% c("TOTAL IMPOR","MIGAS", "NON MIGAS")), paste(" - ", tabel1$Uraian), tabel1$Uraian)
    tabel1$urutan<-NULL
    tabel1$Jenis<-NULL
    tabel1[tabel1 == 0] <- NA
    
    if(bulan!=1){
      tabel1<-mutate(tabel1, tabel1_mtm=100*(!!sym(kolom_bulan_ini)-!!sym(kolom_bulan_lalu))/!!sym(kolom_bulan_lalu), tabel1_yoy=100*(!!sym(kolom_bulan_ini)-!!sym(kolom_tahun_lalu))/!!sym(kolom_tahun_lalu), tabel1_ctc=100*(!!sym(kolom_cumulative)-!!sym(kolom_cumulative_tahun_lalu))/!!sym(kolom_cumulative_tahun_lalu), peran=100*!!sym(kolom_cumulative)/ekspor_cumulative)
      colnames(tabel1)<-c("Uraian",kolom_tahun_lalu, kolom_cumulative_tahun_lalu,kolom_bulan_lalu, kolom_bulan_ini, kolom_cumulative, "m-to-m (%)", "y-on-y (%)", "c-to-c (%)", paste("Peran (%)",kolom_cumulative))
    } else{
      tabel1<-mutate(tabel1, tabel1_mtm=100*(!!sym(kolom_bulan_ini)-!!sym(kolom_bulan_lalu))/!!sym(kolom_bulan_lalu), tabel1_yoy=100*(!!sym(kolom_bulan_ini)-!!sym(kolom_tahun_lalu))/!!sym(kolom_tahun_lalu),peran=100*!!sym(kolom_bulan_ini)/ekspor_bulan_ini)
      colnames(tabel1)<-c("Uraian",kolom_tahun_lalu,kolom_bulan_lalu, kolom_bulan_ini, "m-to-m (%)", "y-on-y (%)", paste("Peran (%)",kolom_bulan_ini))
    }
    
    tabel1 <- tabel1 %>%
      mutate(
        `m-to-m (%)` = ifelse(!is.na(!!sym(kolom_bulan_lalu)) & is.na(!!sym(kolom_bulan_ini)), -100, `m-to-m (%)`),
        `y-on-y (%)` = ifelse(!is.na(!!sym(kolom_tahun_lalu)) & is.na(!!sym(kolom_bulan_ini)), -100, `y-on-y (%)`)
      )
    
    if(bulan != 1){
      tabel1 <- tabel1 %>%
        mutate(
          `c-to-c (%)` = ifelse(!is.na(!!sym(kolom_cumulative_tahun_lalu)) & is.na(!!sym(kolom_cumulative)), -100, `c-to-c (%)`)
        )
    }
    
    
    tabel1_impor<<-tabel1
    
    
    
    ##Tabel Output Nilai Ekspor Nonmigas Menurut Golongan Barang HS2 Digit (Ekspor Menurut HS2 Digit)
    hs2_cumulative_tahun_lalu<-summarise(group_by(data[data$BULAN<=bulan & data$TAHUN==tahun-1 & data$JENIS=="NON MIGAS",], HSDUA,HS2), Total_Ekspor = sum(NILAI_JUTA))
    hs2_cumulative<-summarise(group_by(data[data$BULAN<=bulan & data$TAHUN==tahun & data$JENIS=="NON MIGAS",], HSDUA,HS2), Total_Ekspor = sum(NILAI_JUTA))
    hs2_bulan_ini<-summarise(group_by(data[data$BULAN==bulan & data$TAHUN==tahun & data$JENIS=="NON MIGAS",], HSDUA,HS2), Total_Ekspor = sum(NILAI_JUTA))
    hs2_bulan_lalu<-summarise(group_by(data[data$BULAN==bulan_sebelumnya & data$TAHUN==tahun_sebelumnya & data$JENIS=="NON MIGAS",], HSDUA,HS2), Total_Ekspor = sum(NILAI_JUTA))
    hs2_tahun_lalu<-summarise(group_by(data[data$BULAN==bulan & data$TAHUN==tahun-1 & data$JENIS=="NON MIGAS",], HSDUA,HS2), Total_Ekspor = sum(NILAI_JUTA))
    
    
    if (bulan!=1) {
      gabung1_hs2<-full_join(hs2_bulan_ini,hs2_cumulative,by=c("HSDUA","HS2"))
      gabung2_hs2<-full_join(hs2_bulan_lalu,gabung1_hs2,by=c("HSDUA","HS2"))
      gabung3_hs2<-full_join(hs2_cumulative_tahun_lalu,gabung2_hs2,by=c("HSDUA","HS2"))
      tabel1_hs2<-full_join(hs2_tahun_lalu,gabung3_hs2,by=c("HSDUA","HS2"))
      colnames(tabel1_hs2)<-c("HS2 Digit","Deskripsi",kolom_tahun_lalu, kolom_cumulative_tahun_lalu,kolom_bulan_lalu, kolom_bulan_ini, kolom_cumulative)
      tabel1_1 <- arrange(tabel1_hs2, desc(!!sym(kolom_cumulative)),desc(!!sym(kolom_cumulative_tahun_lalu)))
      
    } else {
      gabung1_hs2<-full_join(hs2_bulan_lalu,hs2_bulan_ini,by=c("HSDUA","HS2"))
      tabel1_1<-full_join(hs2_tahun_lalu,gabung1_hs2,by=c("HSDUA","HS2"))
      colnames(tabel1_1)<-c("HS2 Digit","Deskripsi",kolom_tahun_lalu, kolom_bulan_lalu, kolom_bulan_ini)
      tabel1_1 <- arrange(tabel1_1, desc(!!sym(kolom_bulan_ini)), desc(!!sym(kolom_bulan_lalu)))
    }
    
    
    if(bulan!=1){
      tabel1_1<-mutate(tabel1_1, tabel1_1_mtm=100*(!!sym(kolom_bulan_ini)-!!sym(kolom_bulan_lalu))/!!sym(kolom_bulan_lalu), tabel1_1_yoy=100*(!!sym(kolom_bulan_ini)-!!sym(kolom_tahun_lalu))/!!sym(kolom_tahun_lalu), tabel1_1_ctc=100*(!!sym(kolom_cumulative)-!!sym(kolom_cumulative_tahun_lalu))/!!sym(kolom_cumulative_tahun_lalu), peran=100*!!sym(kolom_cumulative)/ekspor_cumulative_nonmigas)
      colnames(tabel1_1)<-c("Kode HS","Deskripsi",kolom_tahun_lalu, kolom_cumulative_tahun_lalu,kolom_bulan_lalu, kolom_bulan_ini, kolom_cumulative, "m-to-m (%)", "y-on-y (%)", "c-to-c (%)", paste("Peran (%)",kolom_cumulative))
    } else{
      tabel1_1<-mutate(tabel1_1, tabel1_1_mtm=100*(!!sym(kolom_bulan_ini)-!!sym(kolom_bulan_lalu))/!!sym(kolom_bulan_lalu), tabel1_1_yoy=100*(!!sym(kolom_bulan_ini)-!!sym(kolom_tahun_lalu))/!!sym(kolom_tahun_lalu),peran=100*!!sym(kolom_bulan_ini)/ekspor_bulan_ini_nonmigas)
      colnames(tabel1_1)<-c("Kode HS","Deskripsi",kolom_tahun_lalu,kolom_bulan_lalu, kolom_bulan_ini, "m-to-m (%)", "y-on-y (%)", paste("Peran (%)",kolom_bulan_ini))
    }
    
    tabel1_1 <- tabel1_1 %>%
      mutate(
        `m-to-m (%)` = ifelse(!is.na(!!sym(kolom_bulan_lalu)) & is.na(!!sym(kolom_bulan_ini)), -100, `m-to-m (%)`),
        `y-on-y (%)` = ifelse(!is.na(!!sym(kolom_tahun_lalu)) & is.na(!!sym(kolom_bulan_ini)), -100, `y-on-y (%)`)
      )
    
    if(bulan != 1){
      tabel1_1 <- tabel1_1 %>%
        mutate(
          `c-to-c (%)` = ifelse(!is.na(!!sym(kolom_cumulative_tahun_lalu)) & is.na(!!sym(kolom_cumulative)), -100, `c-to-c (%)`)
        )
    }
    
    tabel13_impor<<-tabel1_1
    
    
    
    ##Tabel Output Nilai Ekspor Nonmigas Menurut Negara Tujuan (Ekspor Menurut Negara Tujuan)
    negara_nonmigas_cumulative_tahun_lalu<-summarise(group_by(data[data$BULAN<=bulan & data$TAHUN==tahun-1 & data$JENIS=="NON MIGAS",], NEGARA2), Total_Ekspor = sum(NILAI_JUTA))
    negara_nonmigas_cumulative<-summarise(group_by(data[data$BULAN<=bulan & data$TAHUN==tahun & data$JENIS=="NON MIGAS",], NEGARA2), Total_Ekspor = sum(NILAI_JUTA))
    negara_nonmigas_bulan_ini<-summarise(group_by(data[data$BULAN==bulan & data$TAHUN==tahun & data$JENIS=="NON MIGAS",], NEGARA2), Total_Ekspor = sum(NILAI_JUTA))
    negara_nonmigas_bulan_lalu<-summarise(group_by(data[data$BULAN==bulan_sebelumnya & data$TAHUN==tahun_sebelumnya & data$JENIS=="NON MIGAS",], NEGARA2), Total_Ekspor = sum(NILAI_JUTA))
    negara_nonmigas_tahun_lalu<-summarise(group_by(data[data$BULAN==bulan & data$TAHUN==tahun-1 & data$JENIS=="NON MIGAS",], NEGARA2), Total_Ekspor = sum(NILAI_JUTA))
    
    
    if (bulan!=1) {
      gabung1_negara_nonmigas<-full_join(negara_nonmigas_bulan_ini,negara_nonmigas_cumulative,by="NEGARA2")
      gabung2_negara_nonmigas<-full_join(negara_nonmigas_bulan_lalu,gabung1_negara_nonmigas,by="NEGARA2")
      gabung3_negara_nonmigas<-full_join(negara_nonmigas_cumulative_tahun_lalu,gabung2_negara_nonmigas,by="NEGARA2")
      tabel1_negara_nonmigas<-full_join(negara_nonmigas_tahun_lalu,gabung3_negara_nonmigas,by="NEGARA2")
      colnames(tabel1_negara_nonmigas)<-c("Negara",kolom_tahun_lalu, kolom_cumulative_tahun_lalu,kolom_bulan_lalu, kolom_bulan_ini, kolom_cumulative)
      tabel1_2 <- arrange(tabel1_negara_nonmigas, desc(!!sym(kolom_cumulative)), desc(!!sym(kolom_cumulative_tahun_lalu)))
      
    } else {
      gabung1_negara_nonmigas<-full_join(negara_nonmigas_bulan_lalu,negara_nonmigas_bulan_ini,by="NEGARA2")
      tabel1_2<-full_join(negara_nonmigas_tahun_lalu,gabung1_negara_nonmigas,by="NEGARA2")
      colnames(tabel1_2)<-c("Negara",kolom_tahun_lalu, kolom_bulan_lalu, kolom_bulan_ini)
      tabel1_2 <- arrange(tabel1_2, desc(!!sym(kolom_bulan_ini)), desc(!!sym(kolom_bulan_lalu)))
    }
    
    
    if(bulan!=1){
      tabel1_2<-mutate(tabel1_2, tabel1_2_mtm=100*(!!sym(kolom_bulan_ini)-!!sym(kolom_bulan_lalu))/!!sym(kolom_bulan_lalu), tabel1_2_yoy=100*(!!sym(kolom_bulan_ini)-!!sym(kolom_tahun_lalu))/!!sym(kolom_tahun_lalu), tabel1_2_ctc=100*(!!sym(kolom_cumulative)-!!sym(kolom_cumulative_tahun_lalu))/!!sym(kolom_cumulative_tahun_lalu), peran=100*!!sym(kolom_cumulative)/ekspor_cumulative_nonmigas)
      colnames(tabel1_2)<-c("Negara",kolom_tahun_lalu, kolom_cumulative_tahun_lalu,kolom_bulan_lalu, kolom_bulan_ini, kolom_cumulative, "m-to-m (%)", "y-on-y (%)", "c-to-c (%)", paste("Peran (%)",kolom_cumulative))
    } else{
      tabel1_2<-mutate(tabel1_2, tabel1_2_mtm=100*(!!sym(kolom_bulan_ini)-!!sym(kolom_bulan_lalu))/!!sym(kolom_bulan_lalu), tabel1_2_yoy=100*(!!sym(kolom_bulan_ini)-!!sym(kolom_tahun_lalu))/!!sym(kolom_tahun_lalu),peran=100*!!sym(kolom_bulan_ini)/ekspor_bulan_ini_nonmigas)
      colnames(tabel1_2)<-c("Negara",kolom_tahun_lalu,kolom_bulan_lalu, kolom_bulan_ini, "m-to-m (%)", "y-on-y (%)", paste("Peran (%)",kolom_bulan_ini))
    }
    
    tabel1_2 <- tabel1_2 %>%
      mutate(
        `m-to-m (%)` = ifelse(!is.na(!!sym(kolom_bulan_lalu)) & is.na(!!sym(kolom_bulan_ini)), -100, `m-to-m (%)`),
        `y-on-y (%)` = ifelse(!is.na(!!sym(kolom_tahun_lalu)) & is.na(!!sym(kolom_bulan_ini)), -100, `y-on-y (%)`)
      )
    
    if(bulan != 1){
      tabel1_2 <- tabel1_2 %>%
        mutate(
          `c-to-c (%)` = ifelse(!is.na(!!sym(kolom_cumulative_tahun_lalu)) & is.na(!!sym(kolom_cumulative)), -100, `c-to-c (%)`)
        )
    }
    
    tabel7_impor<<-tabel1_2
    
    
    ##Tabel Output Nilai Ekspor Migas Menurut Negara Tujuan (Ekspor Menurut Negara Tujuan)
    negara_migas_cumulative_tahun_lalu<-summarise(group_by(data[data$BULAN<=bulan & data$TAHUN==tahun-1 & data$JENIS=="MIGAS",], NEGARA2), Total_Ekspor = sum(NILAI_JUTA))
    negara_migas_cumulative<-summarise(group_by(data[data$BULAN<=bulan & data$TAHUN==tahun & data$JENIS=="MIGAS",], NEGARA2), Total_Ekspor = sum(NILAI_JUTA))
    negara_migas_bulan_ini<-summarise(group_by(data[data$BULAN==bulan & data$TAHUN==tahun & data$JENIS=="MIGAS",], NEGARA2), Total_Ekspor = sum(NILAI_JUTA))
    negara_migas_bulan_lalu<-summarise(group_by(data[data$BULAN==bulan_sebelumnya & data$TAHUN==tahun_sebelumnya & data$JENIS=="MIGAS",], NEGARA2), Total_Ekspor = sum(NILAI_JUTA))
    negara_migas_tahun_lalu<-summarise(group_by(data[data$BULAN==bulan & data$TAHUN==tahun-1 & data$JENIS=="MIGAS",], NEGARA2), Total_Ekspor = sum(NILAI_JUTA))
    
    
    if (bulan!=1) {
      gabung1_negara_migas<-full_join(negara_migas_bulan_ini,negara_migas_cumulative,by="NEGARA2")
      gabung2_negara_migas<-full_join(negara_migas_bulan_lalu,gabung1_negara_migas,by="NEGARA2")
      gabung3_negara_migas<-full_join(negara_migas_cumulative_tahun_lalu,gabung2_negara_migas,by="NEGARA2")
      tabel1_negara_migas<-full_join(negara_migas_tahun_lalu,gabung3_negara_migas,by="NEGARA2")
      colnames(tabel1_negara_migas)<-c("Negara",kolom_tahun_lalu, kolom_cumulative_tahun_lalu,kolom_bulan_lalu, kolom_bulan_ini, kolom_cumulative)
      tabel1_3 <- arrange(tabel1_negara_migas, desc(!!sym(kolom_cumulative)),desc(!!sym(kolom_cumulative_tahun_lalu)))
      
    } else {
      gabung1_negara_migas<-full_join(negara_migas_bulan_lalu,negara_migas_bulan_ini,by="NEGARA2")
      tabel1_3<-full_join(negara_migas_tahun_lalu,gabung1_negara_migas,by="NEGARA2")
      colnames(tabel1_3)<-c("Negara",kolom_tahun_lalu, kolom_bulan_lalu, kolom_bulan_ini)
      tabel1_3 <- arrange(tabel1_3, desc(!!sym(kolom_bulan_ini)),desc(!!sym(kolom_bulan_lalu)))
    }
    
    
    if(bulan!=1){
      tabel1_3<-mutate(tabel1_3, tabel1_3_mtm=100*(!!sym(kolom_bulan_ini)-!!sym(kolom_bulan_lalu))/!!sym(kolom_bulan_lalu), tabel1_3_yoy=100*(!!sym(kolom_bulan_ini)-!!sym(kolom_tahun_lalu))/!!sym(kolom_tahun_lalu), tabel1_3_ctc=100*(!!sym(kolom_cumulative)-!!sym(kolom_cumulative_tahun_lalu))/!!sym(kolom_cumulative_tahun_lalu), peran=100*!!sym(kolom_cumulative)/ekspor_cumulative_migas)
      colnames(tabel1_3)<-c("Negara",kolom_tahun_lalu, kolom_cumulative_tahun_lalu,kolom_bulan_lalu, kolom_bulan_ini, kolom_cumulative, "m-to-m (%)", "y-on-y (%)", "c-to-c (%)", paste("Peran (%)",kolom_cumulative))
    } else{
      tabel1_3<-mutate(tabel1_3, tabel1_3_mtm=100*(!!sym(kolom_bulan_ini)-!!sym(kolom_bulan_lalu))/!!sym(kolom_bulan_lalu), tabel1_3_yoy=100*(!!sym(kolom_bulan_ini)-!!sym(kolom_tahun_lalu))/!!sym(kolom_tahun_lalu),peran=100*!!sym(kolom_bulan_ini)/ekspor_bulan_ini_migas)
      colnames(tabel1_3)<-c("Negara",kolom_tahun_lalu,kolom_bulan_lalu, kolom_bulan_ini, "m-to-m (%)", "y-on-y (%)", paste("Peran (%)",kolom_bulan_ini))
    }
    
    tabel1_3 <- tabel1_3 %>%
      mutate(
        `m-to-m (%)` = ifelse(!is.na(!!sym(kolom_bulan_lalu)) & is.na(!!sym(kolom_bulan_ini)), -100, `m-to-m (%)`),
        `y-on-y (%)` = ifelse(!is.na(!!sym(kolom_tahun_lalu)) & is.na(!!sym(kolom_bulan_ini)), -100, `y-on-y (%)`)
      )
    
    if(bulan != 1){
      tabel1_3 <- tabel1_3 %>%
        mutate(
          `c-to-c (%)` = ifelse(!is.na(!!sym(kolom_cumulative_tahun_lalu)) & is.na(!!sym(kolom_cumulative)), -100, `c-to-c (%)`)
        )
    }
    
    tabel8_impor<<-tabel1_3
    
    
    
    ##Tabel Output Nilai Ekspor Menurut Pelabuhan (Ringkasan Nilai/Volume Ekspor)
    pelabuhan_cumulative_tahun_lalu<-summarise(group_by(data[data$BULAN<=bulan & data$TAHUN==tahun-1,], PELABUHAN1), Total_Ekspor = sum(NILAI_JUTA))
    pelabuhan_cumulative<-summarise(group_by(data[data$BULAN<=bulan & data$TAHUN==tahun,], PELABUHAN1), Total_Ekspor = sum(NILAI_JUTA))
    pelabuhan_bulan_ini<-summarise(group_by(data[data$BULAN==bulan & data$TAHUN==tahun,], PELABUHAN1), Total_Ekspor = sum(NILAI_JUTA))
    pelabuhan_bulan_lalu<-summarise(group_by(data[data$BULAN==bulan_sebelumnya & data$TAHUN==tahun_sebelumnya,], PELABUHAN1), Total_Ekspor = sum(NILAI_JUTA))
    pelabuhan_tahun_lalu<-summarise(group_by(data[data$BULAN==bulan & data$TAHUN==tahun-1,], PELABUHAN1), Total_Ekspor = sum(NILAI_JUTA))
    
    
    if (bulan!=1) {
      gabung1_pelabuhan<-full_join(pelabuhan_bulan_ini,pelabuhan_cumulative,by="PELABUHAN1")
      gabung2_pelabuhan<-full_join(pelabuhan_bulan_lalu,gabung1_pelabuhan,by="PELABUHAN1")
      gabung3_pelabuhan<-full_join(pelabuhan_cumulative_tahun_lalu,gabung2_pelabuhan,by="PELABUHAN1")
      tabel1_pelabuhan<-full_join(pelabuhan_tahun_lalu,gabung3_pelabuhan,by="PELABUHAN1")
      colnames(tabel1_pelabuhan)<-c("Pelabuhan",kolom_tahun_lalu, kolom_cumulative_tahun_lalu,kolom_bulan_lalu, kolom_bulan_ini, kolom_cumulative)
      tabel1_4 <- arrange(tabel1_pelabuhan, desc(!!sym(kolom_cumulative)), desc(!!sym(kolom_cumulative_tahun_lalu)))
      
    } else {
      gabung1_pelabuhan<-full_join(pelabuhan_bulan_lalu,pelabuhan_bulan_ini,by="PELABUHAN1")
      tabel1_4<-full_join(pelabuhan_tahun_lalu,gabung1_pelabuhan,by="PELABUHAN1")
      colnames(tabel1_4)<-c("Pelabuhan",kolom_tahun_lalu, kolom_bulan_lalu, kolom_bulan_ini)
      tabel1_4 <- arrange(tabel1_4, desc(!!sym(kolom_bulan_ini)), desc(!!sym(kolom_bulan_lalu)))
    }
    
    
    if(bulan!=1){
      tabel1_4<-mutate(tabel1_4, tabel1_4_mtm=100*(!!sym(kolom_bulan_ini)-!!sym(kolom_bulan_lalu))/!!sym(kolom_bulan_lalu), tabel1_4_yoy=100*(!!sym(kolom_bulan_ini)-!!sym(kolom_tahun_lalu))/!!sym(kolom_tahun_lalu), tabel1_4_ctc=100*(!!sym(kolom_cumulative)-!!sym(kolom_cumulative_tahun_lalu))/!!sym(kolom_cumulative_tahun_lalu), peran=100*!!sym(kolom_cumulative)/ekspor_cumulative)
      colnames(tabel1_4)<-c("Pelabuhan",kolom_tahun_lalu, kolom_cumulative_tahun_lalu,kolom_bulan_lalu, kolom_bulan_ini, kolom_cumulative, "m-to-m (%)", "y-on-y (%)", "c-to-c (%)", paste("Peran (%)",kolom_cumulative))
    } else{
      tabel1_4<-mutate(tabel1_4, tabel1_4_mtm=100*(!!sym(kolom_bulan_ini)-!!sym(kolom_bulan_lalu))/!!sym(kolom_bulan_lalu), tabel1_4_yoy=100*(!!sym(kolom_bulan_ini)-!!sym(kolom_tahun_lalu))/!!sym(kolom_tahun_lalu),peran=100*!!sym(kolom_bulan_ini)/ekspor_bulan_ini)
      colnames(tabel1_4)<-c("Pelabuhan",kolom_tahun_lalu,kolom_bulan_lalu, kolom_bulan_ini, "m-to-m (%)", "y-on-y (%)", paste("Peran (%)",kolom_bulan_ini))
    }
    
    tabel1_4 <- tabel1_4 %>%
      mutate(
        `m-to-m (%)` = ifelse(!is.na(!!sym(kolom_bulan_lalu)) & is.na(!!sym(kolom_bulan_ini)), -100, `m-to-m (%)`),
        `y-on-y (%)` = ifelse(!is.na(!!sym(kolom_tahun_lalu)) & is.na(!!sym(kolom_bulan_ini)), -100, `y-on-y (%)`)
      )
    
    if(bulan != 1){
      tabel1_4 <- tabel1_4 %>%
        mutate(
          `c-to-c (%)` = ifelse(!is.na(!!sym(kolom_cumulative_tahun_lalu)) & is.na(!!sym(kolom_cumulative)), -100, `c-to-c (%)`)
        )
    }
    
    tabel2_impor<<-tabel1_4
    
    
    
    ## Tabel Output Volume Ekspor Menurut Pelabuhan (Ringkasan Nilai/Volume Ekspor)
    
    # Total ekspor (cumulative dan bulan ini)
    ekspor_cumulative_volume <- sum(data$BERAT[data$BULAN <= bulan & data$TAHUN == tahun], na.rm = TRUE)/1000000
    ekspor_bulan_ini_volume <- sum(data$BERAT[data$BULAN == bulan & data$TAHUN == tahun], na.rm = TRUE)/1000000
    
    # Per pelabuhan
    pelabuhan_volume_cumulative_tahun_lalu <- data %>%
      filter(BULAN <= bulan, TAHUN == tahun - 1) %>%
      group_by(PELABUHAN1) %>%
      summarise(Total_Volume = sum(BERAT, na.rm = TRUE) / 1000000)
    
    pelabuhan_volume_cumulative <- data %>%
      filter(BULAN <= bulan, TAHUN == tahun) %>%
      group_by(PELABUHAN1) %>%
      summarise(Total_Volume = sum(BERAT, na.rm = TRUE) / 1000000)
    
    pelabuhan_volume_bulan_ini <- data %>%
      filter(BULAN == bulan, TAHUN == tahun) %>%
      group_by(PELABUHAN1) %>%
      summarise(Total_Volume = sum(BERAT, na.rm = TRUE) / 1000000)
    
    pelabuhan_volume_bulan_lalu <- data %>%
      filter(BULAN == bulan_sebelumnya, TAHUN == tahun_sebelumnya) %>%
      group_by(PELABUHAN1) %>%
      summarise(Total_Volume = sum(BERAT, na.rm = TRUE) / 1000000)
    
    pelabuhan_volume_tahun_lalu <- data %>%
      filter(BULAN == bulan, TAHUN == tahun - 1) %>%
      group_by(PELABUHAN1) %>%
      summarise(Total_Volume = sum(BERAT, na.rm = TRUE) / 1000000)
    
    # Penggabungan dan pengolahan
    if (bulan != 1) {
      gabung1_pelabuhan_volume <- full_join(pelabuhan_volume_bulan_ini, pelabuhan_volume_cumulative, by = "PELABUHAN1")
      gabung2_pelabuhan_volume <- full_join(pelabuhan_volume_bulan_lalu, gabung1_pelabuhan_volume, by = "PELABUHAN1")
      gabung3_pelabuhan_volume <- full_join(pelabuhan_volume_cumulative_tahun_lalu, gabung2_pelabuhan_volume, by = "PELABUHAN1")
      tabel1_pelabuhan_volume <- full_join(pelabuhan_volume_tahun_lalu, gabung3_pelabuhan_volume, by = "PELABUHAN1")
      
      # Penamaan kolom
      colnames(tabel1_pelabuhan_volume) <- c("Pelabuhan", kolom_tahun_lalu, kolom_cumulative_tahun_lalu,
                                             kolom_bulan_lalu, kolom_bulan_ini, kolom_cumulative)
      
      # Urutkan berdasarkan kumulatif
      tabel1_5 <- arrange(tabel1_pelabuhan_volume, desc(!!sym(kolom_cumulative)),desc(!!sym(kolom_cumulative_tahun_lalu)))
      
      # Tambahkan mtm, yoy, ctc, peran
      tabel1_5 <- tabel1_5 %>%
        mutate(
          `m-to-m (%)` = 100 * (!!sym(kolom_bulan_ini) - !!sym(kolom_bulan_lalu)) / !!sym(kolom_bulan_lalu),
          `y-on-y (%)` = 100 * (!!sym(kolom_bulan_ini) - !!sym(kolom_tahun_lalu)) / !!sym(kolom_tahun_lalu),
          `c-to-c (%)` = 100 * (!!sym(kolom_cumulative) - !!sym(kolom_cumulative_tahun_lalu)) / !!sym(kolom_cumulative_tahun_lalu),
          !!paste("Peran (%)", kolom_cumulative) := 100 * !!sym(kolom_cumulative) / ekspor_cumulative_volume
        )
      
    } else {
      gabung1_pelabuhan_volume <- full_join(pelabuhan_volume_bulan_lalu, pelabuhan_volume_bulan_ini, by = "PELABUHAN1")
      tabel1_5 <- full_join(pelabuhan_volume_tahun_lalu, gabung1_pelabuhan_volume, by = "PELABUHAN1")
      
      # Penamaan kolom
      colnames(tabel1_5) <- c("Pelabuhan", kolom_tahun_lalu, kolom_bulan_lalu, kolom_bulan_ini)
      
      # Urutkan berdasarkan bulan ini
      tabel1_5 <- arrange(tabel1_5, desc(!!sym(kolom_bulan_ini)),  desc(!!sym(kolom_bulan_lalu)))
      
      # Tambahkan mtm, yoy, peran
      tabel1_5 <- tabel1_5 %>%
        mutate(
          `m-to-m (%)` = 100 * (!!sym(kolom_bulan_ini) - !!sym(kolom_bulan_lalu)) / !!sym(kolom_bulan_lalu),
          `y-on-y (%)` = 100 * (!!sym(kolom_bulan_ini) - !!sym(kolom_tahun_lalu)) / !!sym(kolom_tahun_lalu),
          !!paste("Peran (%)", kolom_bulan_ini) := 100 * !!sym(kolom_bulan_ini) / ekspor_bulan_ini_volume
        )
    }
    
    tabel1_5 <- tabel1_5 %>%
      mutate(
        `m-to-m (%)` = ifelse(!is.na(!!sym(kolom_bulan_lalu)) & is.na(!!sym(kolom_bulan_ini)), -100, `m-to-m (%)`),
        `y-on-y (%)` = ifelse(!is.na(!!sym(kolom_tahun_lalu)) & is.na(!!sym(kolom_bulan_ini)), -100, `y-on-y (%)`)
      )
    
    if(bulan != 1){
      tabel1_5 <- tabel1_5 %>%
        mutate(
          `c-to-c (%)` = ifelse(!is.na(!!sym(kolom_cumulative_tahun_lalu)) & is.na(!!sym(kolom_cumulative)), -100, `c-to-c (%)`)
        )
    }
    
    # Simpan hasil akhir
    tabel5_impor <<- tabel1_5
    
    
    
    ## ----------------------------------------------------------------- SLIDE 2 ---------------------------------------------------------------------------------------------
    rekap_HS2 <- summarise(group_by(data[data$BULAN==bulan & data$TAHUN==tahun & data$JENIS=="NON MIGAS",], HSDUA, HS2), Total_Ekspor = sum(NILAI_JUTA))
    rekap_HS2<-mutate(rekap_HS2,share=100*Total_Ekspor/ekspor_bulan_ini_nonmigas)
    rekap_HS2<-arrange(rekap_HS2,desc(share))
    tabel10_1<-rekap_HS2
    colnames(tabel10_1)<-c("Kode HS", "Deskripsi", paste("Nilai Impor ", kolom_bulan_ini), "Share (%)")
    tabel15_impor<<-tabel10_1
    
    
    rekap_HS2_bulan_ini <- summarise(group_by(data[data$BULAN==bulan & data$TAHUN==tahun & data$JENIS=="NON MIGAS",], HSDUA, HS2), Total_Ekspor = sum(NILAI_JUTA))
    rekap_HS2_bulan_lalu <- summarise(group_by(data[data$BULAN==bulan_sebelumnya & data$TAHUN==tahun_sebelumnya & data$JENIS=="NON MIGAS",], HSDUA, HS2), Total_Ekspor = sum(NILAI_JUTA))
    
    
    rekap_HS2_gabung <- full_join(rekap_HS2_bulan_lalu, rekap_HS2_bulan_ini, by= "HSDUA")
    rekap_HS2_gabung <- rename(rekap_HS2_gabung, Ekspor_Bulan_Lalu=Total_Ekspor.x, Ekspor_Bulan_Ini=Total_Ekspor.y)
    
    rekap_HS2_gabung$HS2.x <- ifelse(is.na(rekap_HS2_gabung$HS2.x), rekap_HS2_gabung$HS2.y, rekap_HS2_gabung$HS2.x)
    rekap_HS2_gabung$HS2.y <- ifelse(is.na(rekap_HS2_gabung$HS2.y), rekap_HS2_gabung$HS2.x, rekap_HS2_gabung$HS2.y)
    
    rekap_HS2_gabung$Ekspor_Bulan_Lalu <- ifelse(is.na(rekap_HS2_gabung$Ekspor_Bulan_Lalu), 0, rekap_HS2_gabung$Ekspor_Bulan_Lalu)
    rekap_HS2_gabung$Ekspor_Bulan_Ini <- ifelse(is.na(rekap_HS2_gabung$Ekspor_Bulan_Ini), 0, rekap_HS2_gabung$Ekspor_Bulan_Ini)
    
    rekap_HS2_gabung <- select(rekap_HS2_gabung, -HS2.y)
    rekap_HS2_gabung <- rename(rekap_HS2_gabung, HS2=HS2.x)
    
    if ((bulan_sebelumnya %in% data$BULAN) & (tahun_sebelumnya %in% data$TAHUN)) {
    } else {
      rekap_HS2_gabung$Ekspor_Bulan_Lalu <- rep(NA_real_, nrow(rekap_HS2_gabung))
    }
    
    rekap_HS2_gabung <- mutate(rekap_HS2_gabung, peningkatan=Ekspor_Bulan_Ini-Ekspor_Bulan_Lalu)
    tabel9_1<<-arrange(rekap_HS2_gabung,desc(peningkatan))
    colnames(tabel9_1)<-c("Kode HS","Deskripsi",paste("Nilai Impor ", kolom_bulan_lalu),paste("Nilai Impor ", kolom_bulan_ini),"Peningkatan Nilai Impor")
    
    tabel14_impor<<-tabel9_1
    
    ## ----------------------------------------------------------------- SLIDE 3 ---------------------------------------------------------------------------------------------
    
    rekap_negara <- summarise(group_by(data[data$BULAN==bulan & data$TAHUN==tahun,], NEGARA2), Total_Ekspor = sum(NILAI_JUTA))
    rekap_negara <- mutate(rekap_negara, share=100*Total_Ekspor/ekspor_bulan_ini)
    rekap_negara<-arrange(rekap_negara,desc(share))
    tabel4_1<-rekap_negara
    colnames(tabel4_1)<-c("Negara", paste("Nilai Impor ", kolom_bulan_ini), "Share (%)")
    tabel6_impor<<-tabel4_1
    
    
    
    rekap_HS2_negara_utama <- summarise(group_by(data[data$BULAN==bulan & data$TAHUN==tahun & data$NEGARA2==rekap_negara$NEGARA2[1],], HSDUA, HS2), Total_Ekspor = sum(NILAI_JUTA))
    ekspor_bulan_ini_negara_utama <- rekap_negara$Total_Ekspor[1]
    rekap_HS2_negara_utama <- mutate(rekap_HS2_negara_utama, share=100*Total_Ekspor/ekspor_bulan_ini_negara_utama)
    rekap_HS2_negara_utama <- arrange(rekap_HS2_negara_utama,desc(share))
    tabel5_1<-rekap_HS2_negara_utama
    colnames(tabel5_1)<-c("Kode HS","Deskripsi", paste("Nilai Ekspor ", kolom_bulan_ini), "Share (%)")
    tabel9_impor<<-tabel5_1
    
    
    ekspor_negara_utama_waktu <- summarise(group_by(data[data$NEGARA2==rekap_negara$NEGARA2[1],], BULAN, TAHUN), Total_Ekspor = sum(NILAI_JUTA))
    ekspor_negara_utama_waktu <- arrange(ekspor_negara_utama_waktu,TAHUN, BULAN)
    
    
    kumpulan_tahun <- sort(unique(data$TAHUN[data$TAHUN<=tahun]))
    
    tabel_ekspor_negara_utama_waktu<-matrix(nrow = 12,ncol = 1+length(kumpulan_tahun))
    colnames(tabel_ekspor_negara_utama_waktu) <- rep("", ncol(tabel_ekspor_negara_utama_waktu))
    
    for(i in 1:ncol(tabel_ekspor_negara_utama_waktu)){
      if(i==1){colnames(tabel_ekspor_negara_utama_waktu)[i] <- "Bulan"}
      else {colnames(tabel_ekspor_negara_utama_waktu)[i] <- paste("Nilai Impor ", kumpulan_tahun[i-1])}}
    
    tabel_ekspor_negara_utama_waktu<-as.data.frame(tabel_ekspor_negara_utama_waktu)
    tabel_ekspor_negara_utama_waktu$Bulan<-month.name
    
    for(i in 2:ncol(tabel_ekspor_negara_utama_waktu)){
      for(j in 1:12){
        tabel_ekspor_negara_utama_waktu[j,i]<- sum(ekspor_negara_utama_waktu$Total_Ekspor[ekspor_negara_utama_waktu$BULAN==j&ekspor_negara_utama_waktu$TAHUN==kumpulan_tahun[i-1]])
        if(j>bulan && i>=ncol(tabel_ekspor_negara_utama_waktu)){
          tabel_ekspor_negara_utama_waktu[j,i]<- NA}}}
    
    tabel10_impor<<- tabel_ekspor_negara_utama_waktu
    
    rekap_negara_cumulative <- summarise(group_by(data[data$BULAN<=bulan & data$TAHUN==tahun,], NEGARA2), Total_Ekspor = sum(NILAI_JUTA))
    rekap_negara_cumulative <- mutate(rekap_negara_cumulative, share=100*Total_Ekspor/ekspor_cumulative)
    rekap_negara_cumulative<-arrange(rekap_negara_cumulative,desc(share))
    tabel7_1<-rekap_negara_cumulative
    colnames(tabel7_1)<-c("Negara", paste("Nilai Impor ", kolom_cumulative), "Share (%)")
    tabel11_impor<<-tabel7_1
    
    tabel8_1 <- summarise(group_by(data[data$BULAN <= bulan & data$NEGARA2 == rekap_negara$NEGARA2[1] & data$TAHUN<=tahun, ],TAHUN),Total_Ekspor = sum(NILAI_JUTA, na.rm = TRUE))
    tabel8_1<-arrange(tabel8_1, TAHUN)
    tabel8_1$TAHUN<-as.character(tabel8_1$TAHUN)
    colnames(tabel8_1)<-c(paste("Periode ",month.abb[1]," - ",month.abb[bulan]), "Nilai Impor Kumulatif")
    tabel12_impor<<-tabel8_1
    
    ## ----------------------------------------------------------------- SLIDE 4 ---------------------------------------------------------------------------------------------
    perkembangan_ekspor <- summarise(group_by(data, BULAN, TAHUN), Total_Ekspor = sum(NILAI_JUTA))
    perkembangan_ekspor <- arrange(perkembangan_ekspor,TAHUN, BULAN)
    
    
    tabel_perkembangan_ekspor<-matrix(nrow = 12,ncol = 1+length(kumpulan_tahun))
    colnames(tabel_perkembangan_ekspor) <- rep("", ncol(tabel_perkembangan_ekspor))
    
    for(i in 1:ncol(tabel_perkembangan_ekspor)){
      if(i==1){colnames(tabel_perkembangan_ekspor)[i] <- "Bulan"}
      else {colnames(tabel_perkembangan_ekspor)[i] <- paste("Nilai Impor ",kumpulan_tahun[i-1])}}
    
    tabel_perkembangan_ekspor<-as.data.frame(tabel_perkembangan_ekspor)
    tabel_perkembangan_ekspor$Bulan<-month.name
    
    for(i in 2:ncol(tabel_perkembangan_ekspor)){
      for(j in 1:12){
        tabel_perkembangan_ekspor[j,i]<- sum(perkembangan_ekspor$Total_Ekspor[perkembangan_ekspor$BULAN==j&perkembangan_ekspor$TAHUN==kumpulan_tahun[i-1]])
        if(j>bulan && i>=ncol(tabel_perkembangan_ekspor)){
          tabel_perkembangan_ekspor[j,i]<- NA}}}
    
    tabel3_impor<<-tabel_perkembangan_ekspor
    
    
    
    tabel3_1 <- summarise(group_by(data[data$BULAN <= bulan  & data$TAHUN<=tahun, ],TAHUN),Total_Ekspor = sum(NILAI_JUTA, na.rm = TRUE))
    tabel3_1<-arrange(tabel3_1, TAHUN)
    tabel3_1$TAHUN<-as.character(tabel3_1$TAHUN)
    colnames(tabel3_1)<-c(paste("Periode ",month.abb[1]," - ",month.abb[bulan]), "Nilai Impor Kumulatif")
    tabel4_impor<<-tabel3_1
    
    
    # Reaktif data berdasarkan pilihan
    datasetInput <- reactive({
      switch(input$pilih_tabel_impor,
             "1. Nilai Impor Menurut Sektor" = tabel1_impor,
             "2. Nilai Impor Menurut Pelabuhan" = tabel2_impor,
             "3. Perkembangan Nilai Impor" = tabel3_impor,
             "4. Perkembangan Nilai Impor (c-t-c)" = tabel4_impor,
             "5. Volume Impor Menurut Pelabuhan" = tabel5_impor,
             "6. Nilai Impor Menurut Negara Asal" = tabel6_impor,
             "7. Nilai Impor Nonmigas Menurut Negara Asal" = tabel7_impor,
             "8. Nilai Impor Migas Menurut Negara Asal" = tabel8_impor,
             "9. Nilai Impor Negara Asal Utama HS2 Digit" = tabel9_impor,
             "10. Perkembangan Nilai Impor Negara Asal Utama" = tabel10_impor,
             "11. Nilai Impor Kumulatif Menurut Negara Asal" = tabel11_impor,
             "12. Perkembangan Nilai Impor Negara Asal Utama (c-t-c)" = tabel12_impor,
             "13. Nilai Impor Nonmigas Menurut Golongan Barang HS2 Digit" = tabel13_impor,
             "14. Peningkatan/Penurunan Nilai Impor Nonmigas HS2 Digit (m-t-m)" = tabel14_impor,
             "15. Share Nilai Impor Nonmigas HS2 Digit" = tabel15_impor
             
      )
    })
    
    
    output$downloadDataImpor <- downloadHandler(
      filename = function() {
        paste0(input$pilih_tabel_impor, ".xlsx")
      },
      content = function(file) {
        data_to_download <- datasetInput()
        
        if (is.null(data_to_download) || nrow(data_to_download) == 0) {
          data_to_download <- data.frame(Pesan = "Data tidak tersedia", check.names = FALSE)
        } else {
          # ======== PERBAIKAN NAMA KOLOM ========
          clean_colnames <- function(names_vec) {
            names_vec <- gsub("\\s+", " ", names_vec)         # Hapus spasi berlebih
            names_vec <- trimws(names_vec)                    # Trim kiri-kanan
            # Hanya ubah simbol yang berlebihan
            names_vec <- gsub("(?<![a-zA-Z0-9])-\\s*-\\s*(?![a-zA-Z0-9])", " - ", names_vec, perl = TRUE)  # Perbaiki - yang berlebihan
            return(names_vec)
          }
          
          colnames(data_to_download) <- clean_colnames(colnames(data_to_download))
          
          # ======== FORMAT ANGKA 2 DIGIT ========
          numeric_cols <- sapply(data_to_download, is.numeric)
          
          data_to_download[numeric_cols] <- lapply(
            data_to_download[numeric_cols],
            function(x) {
              out <- ifelse(
                is.na(x), "NA",
                ifelse(x == 0, "ZERO", formatC(x, format = "f", digits = 2, big.mark = ".", decimal.mark = ","))
              )
              return(out)
            }
          )
          
          
          # Ganti "NA" dan "ZERO" menjadi "-"
          data_to_download <- data.frame(
            lapply(data_to_download, function(col) {
              col <- as.character(col)
              col[col %in% c("NA", "ZERO")] <- "-"
              return(col)
            }),
            stringsAsFactors = FALSE,
            check.names = FALSE
          )
          
        }
        
        # ======== MEMBUAT FILE EXCEL ========
        wb <- openxlsx::createWorkbook()
        openxlsx::addWorksheet(wb, "Sheet1")
        
        # Style judul
        title_style <- openxlsx::createStyle(
          fontSize = 14,
          textDecoration = "bold",
          halign = "center"
        )
        
        # Style header (center horizontal & vertical, wrap, border)
        header_style <- openxlsx::createStyle(
          textDecoration = "bold",
          wrapText = TRUE,
          halign = "center",
          valign = "center",
          border = "TopBottomLeftRight",
          borderStyle = "thin"
        )
        
        # Style isi data
        data_style <- openxlsx::createStyle(
          border = "TopBottomLeftRight",
          borderStyle = "thin"
        )
        
        # Tulis judul tabel
        # Buat teks judul
        table_title <- paste("Tabel", input$pilih_tabel_impor)
        
        # Style tanpa wrap dan tanpa merge
        title_style <- openxlsx::createStyle(
          textDecoration = "bold",
          halign = "left",
          valign = "center",
          fontSize = 12,
          wrapText = FALSE
        )
        
        # Tulis judul di sel A1 saja
        openxlsx::writeData(wb, sheet = 1, x = table_title, startRow = 1, startCol = 1, colNames = FALSE)
        openxlsx::addStyle(wb, sheet = 1, style = title_style, rows = 1, cols = 1)
        
        
        # Tulis data (mulai baris ke-2)
        openxlsx::writeData(wb, sheet = 1, x = data_to_download, startRow = 2, headerStyle = header_style)
        
        # Tambahkan border ke isi data
        openxlsx::addStyle(
          wb, sheet = 1, style = data_style,
          rows = 3:(nrow(data_to_download) + 2),
          cols = 1:ncol(data_to_download),
          gridExpand = TRUE,
          stack = TRUE
        )
        
        # ======== TAMBAHKAN CATATAN DI BAWAH TABEL ========
        # Ambil waktu saat ini
        timestamp <- format(Sys.time(), "%d %B %Y pukul %H:%M WIB", tz = "Asia/Jakarta")
        
        # Teks catatan
        footer_text <- paste("Sumber: BPS Kabupaten Karimun (data diolah), diakses pada", timestamp)
        
        # Tentukan baris tempat catatan ditulis (baris setelah data terakhir + 2)
        footer_row <- nrow(data_to_download) + 4  # 1 baris kosong setelah tabel
        
        # Tulis catatan di kolom pertama
        openxlsx::writeData(wb, sheet = 1, x = footer_text, startRow = footer_row, startCol = 1, colNames = FALSE)
        
        # Tambahkan style miring & kecil
        footer_style <- openxlsx::createStyle(
          fontSize = 9,
          textDecoration = "italic",
          halign = "left"
        )
        
        openxlsx::addStyle(wb, sheet = 1, style = footer_style, rows = footer_row, cols = 1, gridExpand = TRUE)
        
        
        # Atur lebar kolom
        get_width <- function(col) {
          max_len <- max(nchar(as.character(col)), na.rm = TRUE)
          return(min(max(10, max_len + 2), 30))
        }
        
        col_widths <- sapply(data_to_download, get_width)
        openxlsx::setColWidths(wb, sheet = 1, cols = 1:ncol(data_to_download), widths = col_widths)
        
        # Simpan
        openxlsx::saveWorkbook(wb, file, overwrite = TRUE)
      }
    )
    
    
    ## Download semua tabel
    output$downloadAllImpor <- downloadHandler(
      filename = function() {
        "Semua Tabel Impor.zip"
      },
      content = function(file) {
        # Buat folder sementara
        temp_dir <- tempdir()
        
        title_style <- openxlsx::createStyle(
          fontSize = 14,
          textDecoration = "bold",
          halign = "center"
        )
        
        # Daftar nama file dan tabel beserta judul tabel
        daftar_tabel <- list(
          "1. Nilai Impor Menurut Sektor.xlsx" = list(tabel = tabel1_impor, judul = "Tabel 1. Nilai Impor Menurut Sektor"),
          "2. Nilai Impor Menurut Pelabuhan.xlsx" = list(tabel = tabel2_impor, judul = "Tabel 2. Nilai Impor Menurut Pelabuhan"),
          "3. Perkembangan Nilai Impor.xlsx" = list(tabel = tabel3_impor, judul = "Tabel 3. Perkembangan Nilai Impor"),
          "4. Perkembangan Nilai Impor (c-t-c).xlsx" = list(tabel = tabel4_impor, judul = "Tabel 4. Perkembangan Nilai Impor (c-t-c)"),
          "5. Volume Impor Menurut Pelabuhan.xlsx" = list(tabel = tabel5_impor, judul = "Tabel 5. Volume Impor Menurut Pelabuhan"),
          "6. Nilai Impor Menurut Negara Asal.xlsx" = list(tabel = tabel6_impor, judul = "Tabel 6. Nilai Impor Menurut Negara Asal"),
          "7. Nilai Impor Nonmigas Menurut Negara Asal.xlsx" = list(tabel = tabel7_impor, judul = "Tabel 7. Nilai Impor Nonmigas Menurut Negara Asal"),
          "8. Nilai Impor Migas Menurut Negara Asal.xlsx" = list(tabel = tabel8_impor, judul = "Tabel 8. Nilai Impor Migas Menurut Negara Asal"),
          "9. Nilai Impor Negara Asal Utama HS2 Digit.xlsx" = list(tabel = tabel9_impor, judul = "Tabel 9. Nilai Impor Negara Asal Utama HS2 Digit"),
          "10. Perkembangan Nilai Impor Negara Asal Utama.xlsx" = list(tabel = tabel10_impor, judul = "Tabel 10. Perkembangan Nilai Impor Negara Asal Utama"),
          "11. Nilai Impor Kumulatif Menurut Negara Asal.xlsx" = list(tabel = tabel11_impor, judul = "Tabel 11. Nilai Impor Kumulatif Menurut Negara Asal"),
          "12. Perkembangan Nilai Impor Negara Asal Utama (c-t-c).xlsx" = list(tabel = tabel12_impor, judul = "Tabel 12. Perkembangan Nilai Impor Negara Asal Utama (c-t-c)"),
          "13. Nilai Impor Nonmigas Menurut Golongan Barang HS2 Digit.xlsx" = list(tabel = tabel13_impor, judul = "Tabel 13. Nilai Impor Nonmigas Menurut Golongan Barang HS2 Digit"),
          "14. Peningkatan_Penurunan Nilai Impor Nonmigas HS2 Digit (m-t-m).xlsx" = list(tabel = tabel14_impor, judul = "Tabel 14. Peningkatan/Penurunan Nilai Impor Nonmigas HS2 Digit (m-t-m)"),
          "15. Share Nilai Impor Nonmigas HS2 Digit.xlsx" = list(tabel = tabel15_impor, judul = "Tabel 15. Share Nilai Impor Nonmigas HS2 Digit")
          
        )
        
        clean_colnames <- function(names_vec) {
          names_vec <- gsub("\\s+", " ", names_vec)
          names_vec <- trimws(names_vec)
          names_vec <- gsub("(?<![a-zA-Z0-9])-\\s*-\\s*(?![a-zA-Z0-9])", " - ", names_vec, perl = TRUE)
          return(names_vec)
        }
        
        format_numeric <- function(df) {
          numeric_cols <- sapply(df, is.numeric)
          
          df[numeric_cols] <- lapply(df[numeric_cols], function(x) {
            out <- ifelse(
              is.na(x), "NA",
              ifelse(x == 0, "ZERO", formatC(x, format = "f", digits = 2, big.mark = ".", decimal.mark = ","))
            )
            return(out)
          })
          
          
          df <- data.frame(
            lapply(df, function(col) {
              col <- as.character(col)
              col[col %in% c("NA", "ZERO")] <- "-"
              return(col)
            }),
            stringsAsFactors = FALSE,
            check.names = FALSE
          )
          
          return(df)
        }
        
        
        write_table_to_excel <- function(data, file_path, judul_tabel) {
          colnames(data) <- clean_colnames(colnames(data))
          data <- format_numeric(data)
          
          wb <- openxlsx::createWorkbook()
          openxlsx::addWorksheet(wb, "Sheet1")
          
          header_style <- openxlsx::createStyle(
            textDecoration = "bold",
            wrapText = TRUE,
            halign = "center",
            valign = "center",
            border = "TopBottomLeftRight",
            borderStyle = "thin"
          )
          
          data_style <- openxlsx::createStyle(
            border = "TopBottomLeftRight",
            borderStyle = "thin"
          )
          
          judul_style <- openxlsx::createStyle(
            textDecoration = "bold",
            halign = "left",
            valign = "center",
            fontSize = 12,
            wrapText = FALSE
          )
          
          openxlsx::writeData(wb, sheet = 1, x = judul_tabel, startRow = 1, startCol = 1, colNames = FALSE)
          openxlsx::addStyle(wb, sheet = 1, style = judul_style, rows = 1, cols = 1)
          
          openxlsx::writeData(wb, sheet = 1, x = data, startRow = 2, headerStyle = header_style)
          openxlsx::addStyle(
            wb, sheet = 1, style = data_style,
            rows = 3:(nrow(data) + 2),
            cols = 1:ncol(data),
            gridExpand = TRUE,
            stack = TRUE
          )
          
          # ===== Tambahkan footer =====
          timestamp <- format(Sys.time(), "%d %B %Y pukul %H:%M WIB", tz = "Asia/Jakarta")
          footer_text <- paste("Sumber: BPS Kabupaten Karimun (data diolah), diakses pada", timestamp)
          footer_row <- nrow(data) + 4
          
          openxlsx::writeData(wb, sheet = 1, x = footer_text, startRow = footer_row, startCol = 1, colNames = FALSE)
          footer_style <- openxlsx::createStyle(
            fontSize = 9,
            textDecoration = "italic",
            halign = "left"
          )
          openxlsx::addStyle(wb, sheet = 1, style = footer_style, rows = footer_row, cols = 1)
          
          get_width <- function(col) {
            max_len <- max(nchar(as.character(col)), na.rm = TRUE)
            return(min(max(10, max_len + 2), 30))
          }
          col_widths <- sapply(data, get_width)
          openxlsx::setColWidths(wb, sheet = 1, cols = 1:ncol(data), widths = col_widths)
          
          openxlsx::saveWorkbook(wb, file_path, overwrite = TRUE)
        }
        
        file_paths <- c()
        for (nama_file in names(daftar_tabel)) {
          path_file <- file.path(temp_dir, nama_file)
          write_table_to_excel(daftar_tabel[[nama_file]]$tabel, path_file, daftar_tabel[[nama_file]]$judul)
          file_paths <- c(file_paths, path_file)
        }
        
        zip::zipr(zipfile = file, files = file_paths)
      },
      contentType = "application/zip"  # <--- INI PENTING!
    )
    
    
    
    output$tabel_data_impor <- renderDT({
      req(data)  # Pastikan data tersedia
      
      data_sorted <- data %>%
        arrange(desc(TAHUN), desc(BULAN))%>%
        select(-NILAI_JUTA)
      
      datatable(
        data_sorted,
        options = list(
          scrollX = TRUE,
          scrollY = "300px",
          fixedHeader = TRUE,
          ordering = TRUE,
          columnDefs = list(
            list(targets = 0, className = "dt-nowrap")
          )
        ),
        rownames = FALSE
      )
    })
    
    
    
    output$tabel_output_impor_3 <- renderDT({
      
      data_terpilih <- switch(input$pilih_tabel_impor,
                              "1. Nilai Impor Menurut Sektor" = tabel1_impor,
                              "2. Nilai Impor Menurut Pelabuhan" = tabel2_impor,
                              "3. Perkembangan Nilai Impor" = tabel3_impor,
                              "4. Perkembangan Nilai Impor (c-t-c)" = tabel4_impor,
                              "5. Volume Impor Menurut Pelabuhan" = tabel5_impor,
                              "6. Nilai Impor Menurut Negara Asal" = tabel6_impor,
                              "7. Nilai Impor Nonmigas Menurut Negara Asal" = tabel7_impor,
                              "8. Nilai Impor Migas Menurut Negara Asal" = tabel8_impor,
                              "9. Nilai Impor Negara Asal Utama HS2 Digit" = tabel9_impor,
                              "10. Perkembangan Nilai Impor Negara Asal Utama" = tabel10_impor,
                              "11. Nilai Impor Kumulatif Menurut Negara Asal" = tabel11_impor,
                              "12. Perkembangan Nilai Impor Negara Asal Utama (c-t-c)" = tabel12_impor,
                              "13. Nilai Impor Nonmigas Menurut Golongan Barang HS2 Digit" = tabel13_impor,
                              "14. Peningkatan/Penurunan Nilai Impor Nonmigas HS2 Digit (m-t-m)" = tabel14_impor,
                              "15. Share Nilai Impor Nonmigas HS2 Digit" = tabel15_impor
                              
      )
      
      numerik_cols <- sapply(data_terpilih, is.numeric)
      
      data_terpilih[numerik_cols] <- lapply(
        data_terpilih[numerik_cols],
        function(x) {
          out <- ifelse(
            is.na(x), "NA",
            ifelse(x == 0, "ZERO", formatC(x, format = "f", digits = 2, big.mark = ".", decimal.mark = ","))
          )
          return(out)
        }
      )
      
      # Ganti "NA" dan "ZERO" jadi "-"
      data_terpilih <- data.frame(
        lapply(data_terpilih, function(col) {
          col <- as.character(col)
          col[col %in% c("NA", "ZERO")] <- "-"
          return(col)
        }),
        stringsAsFactors = FALSE,
        check.names = FALSE
      )
      
      
      
      datatable(
        data_terpilih,
        options = list(
          scrollX = TRUE,
          scrollY = "300px",
          scrollCollapse = TRUE,  # Agar tinggi menyesuaikan jika baris sedikit
          fixedHeader = TRUE,
          ordering = FALSE,
          columnDefs = list(
            list(targets = 0, className = "dt-nowrap")  # Hindari wrap teks kolom pertama
          )
        ),
        rownames = FALSE
      )
      
      
    })
    
  })
  
  
  
  
  ##---------------------------------------------------------Neraca Perdagangan-------------------------------------------
  
  # Reactive values untuk menyimpan data ekspor dan impor
  data_reaktif_neraca_ekspor <- reactiveVal(NULL)
  data_reaktif_neraca_impor  <- reactiveVal(NULL)
  
  # Fungsi pembantu: validasi kolom
  validasi_kolom_tahun_bulan <- function(df) {
    all(c("TAHUN", "BULAN", "NILAI", "BERAT", "PELABUHAN",
          "HS", "NEGARA", "JENIS", "SEKTOR", "KOMODITI",
          "HS2", "HSDUA", "NEGARA2", "PELABUHAN1") %in% names(df))
  }
  
  
  # Observer saat file ekspor diupload
  observeEvent(input$file_input_neraca_ekspor, {
    req(input$file_input_neraca_ekspor)
    
    df <- as.data.frame(haven::read_sav(input$file_input_neraca_ekspor$datapath))
    
    if (!validasi_kolom_tahun_bulan(df)) {
      showNotification("Ada kolom penting yang tidak ditemukan di data ekspor", type = "error")
      return()
    }
    
    data_reaktif_neraca_ekspor(df)
    
    # Jika impor juga sudah diupload, update pilihan tahun
    if (!is.null(data_reaktif_neraca_impor())) {
      update_tahun_dari_kedua_data()
    }
  })
  
  # Observer saat file impor diupload
  observeEvent(input$file_input_neraca_impor, {
    req(input$file_input_neraca_impor)
    
    df <- as.data.frame(haven::read_sav(input$file_input_neraca_impor$datapath))
    
    if (!validasi_kolom_tahun_bulan(df)) {
      showNotification("Ada kolom penting yang tidak ditemukan di data impor", type = "error")
      return()
    }
    
    data_reaktif_neraca_impor(df)
    
    # Jika ekspor juga sudah diupload, update pilihan tahun
    if (!is.null(data_reaktif_neraca_ekspor())) {
      update_tahun_dari_kedua_data()
    }
  })
  
  # Fungsi untuk update pilihan tahun berdasarkan irisannya
  update_tahun_dari_kedua_data <- function() {
    df_ekspor <- data_reaktif_neraca_ekspor()
    df_impor  <- data_reaktif_neraca_impor()
    
    tahun_ekspor <- unique(df_ekspor$TAHUN)
    tahun_impor  <- unique(df_impor$TAHUN)
    
    tahun_iris   <- sort(intersect(tahun_ekspor, tahun_impor))
    
    updateSelectInput(session, "tahun_neraca", choices = tahun_iris)
    updateSelectInput(session, "bulan_neraca", choices = NULL)
  }
  
  # Observer saat tahun dipilih, update pilihan bulan berdasarkan irisan ekspor dan impor
  observeEvent(input$tahun_neraca, {
    req(input$tahun_neraca)
    req(data_reaktif_neraca_ekspor(), data_reaktif_neraca_impor())
    
    df_ekspor <- data_reaktif_neraca_ekspor()
    df_impor  <- data_reaktif_neraca_impor()
    
    bulan_ekspor <- df_ekspor %>%
      filter(TAHUN == input$tahun_neraca) %>%
      pull(BULAN) %>%
      as.numeric()
    
    bulan_impor <- df_impor %>%
      filter(TAHUN == input$tahun_neraca) %>%
      pull(BULAN) %>%
      as.numeric()
    
    # Ambil irisan bulan dari ekspor dan impor
    bulan_angka <- sort(unique(intersect(bulan_ekspor, bulan_impor)))
    bulan_angka <- bulan_angka[bulan_angka %in% 1:12]
    
    # Konversi angka ke nama bulan
    nama_bulan <- month.name[bulan_angka]
    names(bulan_angka) <- nama_bulan
    
    updateSelectInput(session, "bulan_neraca", choices = bulan_angka)
  })
  
  
  
  observeEvent(input$analisis_button_neraca, {
    
    req(data_reaktif_neraca_ekspor(), data_reaktif_neraca_impor())
    req(input$tahun_neraca)
    req(input$bulan_neraca)
    
    tahun_neraca <- as.numeric(input$tahun_neraca)
    bulan_neraca <- as.numeric(input$bulan_neraca)
    
    
    data_ekspor <- data_reaktif_neraca_ekspor()
    data_impor <- data_reaktif_neraca_impor()
    
    
    data_ekspor$BULAN<-as.numeric(data_ekspor$BULAN)
    data_ekspor$TAHUN<-as.numeric(data_ekspor$TAHUN)
    
    data_ekspor$PROPINSI<-as.character(data_ekspor$PROPINSI)
    data_ekspor$PROPINSI <- trimws(data_ekspor$PROPINSI, which = "both")
    
    data_ekspor$PELABUHAN<-as.character(data_ekspor$PELABUHAN)
    data_ekspor$PELABUHAN <- trimws(data_ekspor$PELABUHAN, which = "both")
    
    data_ekspor$HS<-as.character(data_ekspor$HS)
    data_ekspor$HS<-ifelse(nchar(data_ekspor$HS)==7, paste0("0",data_ekspor$HS), data_ekspor$HS)
    data_ekspor$HS <- trimws(data_ekspor$HS, which = "both")
    
    data_ekspor$NEGARA<-as.character(data_ekspor$NEGARA)
    data_ekspor$NEGARA <- trimws(data_ekspor$NEGARA, which = "both")
    
    data_ekspor$BERAT<-as.numeric(data_ekspor$BERAT)
    data_ekspor$NILAI<-as.numeric(data_ekspor$NILAI)
    
    data_ekspor$HSDUA<-as.character(data_ekspor$HSDUA)
    data_ekspor$HSDUA<-ifelse(nchar(data_ekspor$HSDUA)==1, paste0("0",data_ekspor$HSDUA), data_ekspor$HSDUA)
    data_ekspor$HSDUA <- trimws(data_ekspor$HSDUA, which = "both")
    
    data_ekspor$KABKOT<-as.character(data_ekspor$KABKOT)
    data_ekspor$KABKOT <- trimws(data_ekspor$KABKOT, which = "both")
    
    data_ekspor$PELABUHAN1<-as.character(data_ekspor$PELABUHAN1)
    data_ekspor$PELABUHAN1 <- trimws(data_ekspor$PELABUHAN1, which = "both")
    
    data_ekspor$NEGARA2<-as.character(data_ekspor$NEGARA2)
    data_ekspor$NEGARA2 <- trimws(data_ekspor$NEGARA2, which = "both")
    
    data_ekspor$HS2<-as.character(data_ekspor$HS2)
    data_ekspor$HS2 <- trimws(data_ekspor$HS2, which = "both")
    
    data_ekspor$JENIS<-as.character(data_ekspor$JENIS)
    data_ekspor$JENIS <- trimws(data_ekspor$JENIS, which = "both")
    
    data_ekspor$SEKTOR<-as.character(data_ekspor$SEKTOR)
    data_ekspor$SEKTOR <- trimws(data_ekspor$SEKTOR, which = "both")
    
    data_ekspor$KOMODITI<-as.character(data_ekspor$KOMODITI)
    data_ekspor$KOMODITI <- trimws(data_ekspor$KOMODITI, which = "both")
    
    data_ekspor<-mutate(data_ekspor,NILAI_JUTA=NILAI/1000000)
    ##data <- data[data$PROVASAL == "21", ]
    data_ekspor$NILAI_JUTA<-as.numeric(data_ekspor$NILAI_JUTA)
    
    
    data_impor$BULAN<-as.numeric(data_impor$BULAN)
    data_impor$TAHUN<-as.numeric(data_impor$TAHUN)
    
    data_impor$PROPINSI<-as.character(data_impor$PROPINSI)
    data_impor$PROPINSI <- trimws(data_impor$PROPINSI, which = "both")
    
    data_impor$PELABUHAN<-as.character(data_impor$PELABUHAN)
    data_impor$PELABUHAN <- trimws(data_impor$PELABUHAN, which = "both")
    
    data_impor$HS<-as.character(data_impor$HS)
    data_impor$HS<-ifelse(nchar(data_impor$HS)==7, paste0("0",data_impor$HS), data_impor$HS)
    data_impor$HS <- trimws(data_impor$HS, which = "both")
    
    data_impor$NEGARA<-as.character(data_impor$NEGARA)
    data_impor$NEGARA <- trimws(data_impor$NEGARA, which = "both")
    
    data_impor$BERAT<-as.numeric(data_impor$BERAT)
    data_impor$NILAI<-as.numeric(data_impor$NILAI)
    
    data_impor$HSDUA<-as.character(data_impor$HSDUA)
    data_impor$HSDUA<-ifelse(nchar(data_impor$HSDUA)==1, paste0("0",data_impor$HSDUA), data_impor$HSDUA)
    data_impor$HSDUA <- trimws(data_impor$HSDUA, which = "both")
    
    data_impor$KABKOT<-as.character(data_impor$KABKOT)
    data_impor$KABKOT <- trimws(data_impor$KABKOT, which = "both")
    
    data_impor$PELABUHAN1<-as.character(data_impor$PELABUHAN1)
    data_impor$PELABUHAN1 <- trimws(data_impor$PELABUHAN1, which = "both")
    
    data_impor$NEGARA2<-as.character(data_impor$NEGARA2)
    data_impor$NEGARA2 <- trimws(data_impor$NEGARA2, which = "both")
    
    data_impor$HS2<-as.character(data_impor$HS2)
    data_impor$HS2 <- trimws(data_impor$HS2, which = "both")
    
    data_impor$JENIS<-as.character(data_impor$JENIS)
    data_impor$JENIS <- trimws(data_impor$JENIS, which = "both")
    
    data_impor$SEKTOR<-as.character(data_impor$SEKTOR)
    data_impor$SEKTOR <- trimws(data_impor$SEKTOR, which = "both")
    
    data_impor$KOMODITI<-as.character(data_impor$KOMODITI)
    data_impor$KOMODITI <- trimws(data_impor$KOMODITI, which = "both")
    
    data_impor<-mutate(data_impor,NILAI_JUTA=NILAI/1000000)
    ##data <- data[data$PROVASAL == "21", ]
    data_impor$NILAI_JUTA<-as.numeric(data_impor$NILAI_JUTA)
    
    
    ##---------------------------Pembentukan Tabel Neraca-------------------------------------------------
    neraca_ekspor_bulan_ini <- sum(data_ekspor$NILAI_JUTA[as.numeric(data_ekspor$BULAN)==bulan_neraca & data_ekspor$TAHUN==tahun_neraca])
    neraca_impor_bulan_ini <- sum(data_impor$NILAI_JUTA[as.numeric(data_impor$BULAN)==bulan_neraca & data_impor$TAHUN==tahun_neraca])
    neraca_bulan_ini <- neraca_ekspor_bulan_ini - neraca_impor_bulan_ini
    
    neraca_ekspor_bulan_ini_migas <- sum(data_ekspor$NILAI_JUTA[as.numeric(data_ekspor$BULAN)==bulan_neraca & data_ekspor$TAHUN==tahun_neraca & data_ekspor$JENIS=="MIGAS"])
    neraca_impor_bulan_ini_migas <- sum(data_impor$NILAI_JUTA[as.numeric(data_impor$BULAN)==bulan_neraca & data_impor$TAHUN==tahun_neraca & data_impor$JENIS=="MIGAS"])
    neraca_bulan_ini_migas <- neraca_ekspor_bulan_ini_migas - neraca_impor_bulan_ini_migas
    
    neraca_ekspor_bulan_ini_nonmigas <- sum(data_ekspor$NILAI_JUTA[as.numeric(data_ekspor$BULAN)==bulan_neraca & data_ekspor$TAHUN==tahun_neraca & data_ekspor$JENIS=="NON MIGAS"])
    neraca_impor_bulan_ini_nonmigas <- sum(data_impor$NILAI_JUTA[as.numeric(data_impor$BULAN)==bulan_neraca & data_impor$TAHUN==tahun_neraca & data_impor$JENIS=="NON MIGAS"])
    neraca_bulan_ini_nonmigas <- neraca_ekspor_bulan_ini_nonmigas - neraca_impor_bulan_ini_nonmigas
    
    kolom_bulan_ini <- paste(month.abb[bulan_neraca]," ", tahun_neraca)
    
    neraca_ekspor_cumulative <- sum(data_ekspor$NILAI_JUTA[as.numeric(data_ekspor$BULAN)<=bulan_neraca & data_ekspor$TAHUN==tahun_neraca])
    neraca_impor_cumulative <- sum(data_impor$NILAI_JUTA[as.numeric(data_impor$BULAN)<=bulan_neraca & data_impor$TAHUN==tahun_neraca])
    neraca_cumulative <- neraca_ekspor_cumulative - neraca_impor_cumulative
    
    neraca_ekspor_cumulative_migas <- sum(data_ekspor$NILAI_JUTA[as.numeric(data_ekspor$BULAN)<=bulan_neraca & data_ekspor$TAHUN==tahun_neraca & data_ekspor$JENIS=="MIGAS"])
    neraca_impor_cumulative_migas <- sum(data_impor$NILAI_JUTA[as.numeric(data_impor$BULAN)<=bulan_neraca & data_impor$TAHUN==tahun_neraca & data_impor$JENIS=="MIGAS"])
    neraca_cumulative_migas <- neraca_ekspor_cumulative_migas - neraca_impor_cumulative_migas
    
    neraca_ekspor_cumulative_nonmigas <- sum(data_ekspor$NILAI_JUTA[as.numeric(data_ekspor$BULAN)<=bulan_neraca & data_ekspor$TAHUN==tahun_neraca & data_ekspor$JENIS=="NON MIGAS"])
    neraca_impor_cumulative_nonmigas <- sum(data_impor$NILAI_JUTA[as.numeric(data_impor$BULAN)<=bulan_neraca & data_impor$TAHUN==tahun_neraca & data_impor$JENIS=="NON MIGAS"])
    neraca_cumulative_nonmigas <- neraca_ekspor_cumulative_nonmigas - neraca_impor_cumulative_nonmigas
    
    kolom_cumulative <- paste(month.abb[1],"-", month.abb[bulan_neraca]," ", tahun_neraca)
    
    if(bulan_neraca!=1){
      tabel1_1_neraca <- data.frame('Uraian'=c(kolom_bulan_ini,"Total Ekspor/Impor"," - Migas", " - Non Migas"),
                                    'Ekspor'=c(NA,neraca_ekspor_bulan_ini,neraca_ekspor_bulan_ini_migas,neraca_ekspor_bulan_ini_nonmigas),
                                    'Impor'=c(NA,neraca_impor_bulan_ini,neraca_impor_bulan_ini_migas,neraca_impor_bulan_ini_nonmigas),
                                    'Neraca Perdagangan'=c(NA,neraca_bulan_ini,neraca_bulan_ini_migas,neraca_bulan_ini_nonmigas),
                                    check.names = FALSE)
      
      tabel1_2_neraca <- data.frame('Uraian'=c(kolom_cumulative,"Total Ekspor/Impor"," - Migas", " - Non Migas"),
                                    'Ekspor'=c(NA,neraca_ekspor_cumulative,neraca_ekspor_cumulative_migas,neraca_ekspor_cumulative_nonmigas),
                                    'Impor'=c(NA,neraca_impor_cumulative,neraca_impor_cumulative_migas,neraca_impor_cumulative_nonmigas),
                                    'Neraca Perdagangan'=c(NA,neraca_cumulative,neraca_cumulative_migas,neraca_cumulative_nonmigas),
                                    check.names = FALSE)
      
      tabel1_neraca <<- rbind(tabel1_1_neraca, tabel1_2_neraca)
    }else{
      tabel1_neraca <<- data.frame('Uraian'=c(kolom_bulan_ini,"Total Ekspor/Impor"," - Migas", " - Non Migas"),
                                   'Ekspor'=c(NA,neraca_ekspor_bulan_ini,neraca_ekspor_bulan_ini_migas,neraca_ekspor_bulan_ini_nonmigas),
                                   'Impor'=c(NA,neraca_impor_bulan_ini,neraca_impor_bulan_ini_migas,neraca_impor_bulan_ini_nonmigas),
                                   'Neraca Perdagangan'=c(NA,neraca_bulan_ini,neraca_bulan_ini_migas,neraca_bulan_ini_nonmigas),
                                   check.names = FALSE)
    }
    
    
    
    
    # Hitung total ekspor per bulan dan tahun
    perkembangan_neraca_ekspor <- summarise(group_by(data_ekspor, BULAN, TAHUN), Total_Ekspor = sum(NILAI_JUTA))
    perkembangan_neraca_ekspor <- arrange(perkembangan_neraca_ekspor, TAHUN, BULAN)
    
    # Ambil kumpulan tahun yang sesuai
    kumpulan_tahun <- sort(unique(data_ekspor$TAHUN[data_ekspor$TAHUN <= tahun_neraca]))
    
    # Buat matriks untuk ekspor
    tabel_perkembangan_neraca_ekspor <- matrix(nrow = 12, ncol = 1 + length(kumpulan_tahun))
    colnames(tabel_perkembangan_neraca_ekspor) <- rep("", ncol(tabel_perkembangan_neraca_ekspor))
    
    for(i in 1:ncol(tabel_perkembangan_neraca_ekspor)){
      if(i == 1) {
        colnames(tabel_perkembangan_neraca_ekspor)[i] <- "Bulan"
      } else {
        colnames(tabel_perkembangan_neraca_ekspor)[i] <- paste("Nilai Ekspor", kumpulan_tahun[i - 1])
      }
    }
    
    tabel_perkembangan_neraca_ekspor <- as.data.frame(tabel_perkembangan_neraca_ekspor)
    tabel_perkembangan_neraca_ekspor$Bulan <- month.name
    
    for(i in 2:ncol(tabel_perkembangan_neraca_ekspor)){
      for(j in 1:12){
        val <- sum(perkembangan_neraca_ekspor$Total_Ekspor[
          perkembangan_neraca_ekspor$BULAN == j &
            perkembangan_neraca_ekspor$TAHUN == kumpulan_tahun[i - 1]
        ])
        if(j > bulan_neraca && i == ncol(tabel_perkembangan_neraca_ekspor)){
          val <- NA
        }
        tabel_perkembangan_neraca_ekspor[j, i] <- val
      }
    }
    
    # Hitung total impor per bulan dan tahun
    perkembangan_neraca_impor <- summarise(group_by(data_impor, BULAN, TAHUN), Total_Impor = sum(NILAI_JUTA))
    perkembangan_neraca_impor <- arrange(perkembangan_neraca_impor, TAHUN, BULAN)
    
    # Buat matriks untuk impor
    tabel_perkembangan_neraca_impor <- matrix(nrow = 12, ncol = 1 + length(kumpulan_tahun))
    colnames(tabel_perkembangan_neraca_impor) <- rep("", ncol(tabel_perkembangan_neraca_impor))
    
    for(i in 1:ncol(tabel_perkembangan_neraca_impor)){
      if(i == 1) {
        colnames(tabel_perkembangan_neraca_impor)[i] <- "Bulan"
      } else {
        colnames(tabel_perkembangan_neraca_impor)[i] <- paste("Nilai Impor", kumpulan_tahun[i - 1])
      }
    }
    
    tabel_perkembangan_neraca_impor <- as.data.frame(tabel_perkembangan_neraca_impor)
    tabel_perkembangan_neraca_impor$Bulan <- month.name
    
    for(i in 2:ncol(tabel_perkembangan_neraca_impor)){
      for(j in 1:12){
        val <- sum(perkembangan_neraca_impor$Total_Impor[
          perkembangan_neraca_impor$BULAN == j &
            perkembangan_neraca_impor$TAHUN == kumpulan_tahun[i - 1]
        ])
        if(j > bulan_neraca && i == ncol(tabel_perkembangan_neraca_impor)){
          val <- NA
        }
        tabel_perkembangan_neraca_impor[j, i] <- val
      }
    }
    
    # Gabungkan kedua tabel, tanpa duplikasi kolom Bulan
    tabell_neraca <- cbind(
      Bulan = tabel_perkembangan_neraca_ekspor$Bulan,
      tabel_perkembangan_neraca_ekspor[, -1],
      tabel_perkembangan_neraca_impor[, -1]
    )
    
    # Buat tabel neraca (ekspor - impor) per bulan dan tahun
    tabel_neraca <- data.frame(Bulan = month.name)
    for(i in seq_along(kumpulan_tahun)){
      ekspor_col <- 1 + i
      impor_col <- 1 + length(kumpulan_tahun) + i
      tabel_neraca[[paste("Nilai Neraca Perdagangan", kumpulan_tahun[i])]] <- as.numeric(tabell_neraca[, ekspor_col]) - as.numeric(tabell_neraca[, impor_col])
    }
    
    # Simpan tabel neraca perdagangan hasil akhir
    tabel2_neraca <<- tabel_neraca
    
    
    
    
    # Ambil nama kolom neraca dari tabel2_neraca, kecuali kolom Bulan
    kolom_neraca <- colnames(tabel2_neraca)[-1]
    
    # Extract tahun dari nama kolom
    tahun_kolom <- as.numeric(gsub("Nilai Neraca Perdagangan", "", kolom_neraca))
    
    # Filter tahun yang akan dipakai
    tahun_terpakai <- tahun_kolom[tahun_kolom <= tahun_neraca]
    
    # Inisialisasi data frame kosong untuk hasil kumulatif
    hasil_kumulatif <- data.frame(
      Tahun = as.character(tahun_terpakai), 
      Nilai_Neraca_Kumulatif = NA_real_
    )
    
    # Loop per tahun untuk hitung total kumulatif dari bulan 1 sampai bulan_neraca
    for(i in seq_along(tahun_terpakai)) {
      t <- tahun_terpakai[i]
      nama_kolom <- paste0("Nilai Neraca Perdagangan ", t)
      
      # Jika kolom tidak ada, isi NA dan lanjutkan
      if (!(nama_kolom %in% colnames(tabel2_neraca))) {
        hasil_kumulatif$Nilai_Neraca_Kumulatif[i] <- NA
        next
      }
      
      # Ambil nilai neraca dari bulan 1 sampai bulan_neraca untuk tahun t
      nilai_per_bulan <- tabel2_neraca[1:bulan_neraca, nama_kolom]
      
      # Hitung jumlah kumulatif
      hasil_kumulatif$Nilai_Neraca_Kumulatif[i] <- sum(as.numeric(nilai_per_bulan))
    }
    
    # Ganti nama kolom agar lebih deskriptif
    colnames(hasil_kumulatif) <- c(
      paste0("Periode Jan - ", month.abb[bulan_neraca]),
      "Nilai Neraca Perdagangan Kumulatif"
    )
    
    # Simpan ke global environment kalau perlu
    tabel3_neraca <<- hasil_kumulatif
    
    
    
    
    
    # Reaktif data berdasarkan pilihan
    datasetInput <- reactive({
      switch(input$pilih_tabel_neraca,
             "1. Nilai Neraca Perdagangan"=tabel1_neraca,
             "2. Perkembangan Nilai Neraca Perdagangan"=tabel2_neraca,
             "3. Perkembangan Nilai Neraca Perdagangan (c-t-c)"=tabel3_neraca
      )
    })
    
    
    output$downloadDataNeraca <- downloadHandler(
      filename = function() {
        paste0(input$pilih_tabel_neraca, ".xlsx")
      },
      content = function(file) {
        data_to_download <- datasetInput()
        
        if (is.null(data_to_download) || nrow(data_to_download) == 0) {
          data_to_download <- data.frame(Pesan = "Data tidak tersedia", check.names = FALSE)
        } else {
          # ======== PERBAIKAN NAMA KOLOM ========
          clean_colnames <- function(names_vec) {
            names_vec <- gsub("\\s+", " ", names_vec)         # Hapus spasi berlebih
            names_vec <- trimws(names_vec)                    # Trim kiri-kanan
            # Hanya ubah simbol yang berlebihan
            names_vec <- gsub("(?<![a-zA-Z0-9])-\\s*-\\s*(?![a-zA-Z0-9])", " - ", names_vec, perl = TRUE)  # Perbaiki - yang berlebihan
            return(names_vec)
          }
          
          colnames(data_to_download) <- clean_colnames(colnames(data_to_download))
          
          # ======== FORMAT ANGKA 2 DIGIT ========
          numeric_cols <- sapply(data_to_download, is.numeric)
          data_to_download[numeric_cols] <- lapply(data_to_download[numeric_cols], function(x) {
            out <- formatC(x, format = "f", digits = 2, big.mark = ".", decimal.mark = ",")
            out[is.na(x)] <- ""
            out
          })
          
          
          
        }
        
        # ======== MEMBUAT FILE EXCEL ========
        wb <- openxlsx::createWorkbook()
        openxlsx::addWorksheet(wb, "Sheet1")
        
        # Style judul
        title_style <- openxlsx::createStyle(
          fontSize = 14,
          textDecoration = "bold",
          halign = "center"
        )
        
        # Style header (center horizontal & vertical, wrap, border)
        header_style <- openxlsx::createStyle(
          textDecoration = "bold",
          wrapText = TRUE,
          halign = "center",
          valign = "center",
          border = "TopBottomLeftRight",
          borderStyle = "thin"
        )
        
        # Style isi data
        data_style <- openxlsx::createStyle(
          border = "TopBottomLeftRight",
          borderStyle = "thin"
        )
        
        # Tulis judul tabel
        # Buat teks judul
        table_title <- paste("Tabel", input$pilih_tabel_neraca)
        
        # Style tanpa wrap dan tanpa merge
        title_style <- openxlsx::createStyle(
          textDecoration = "bold",
          halign = "left",
          valign = "center",
          fontSize = 12,
          wrapText = FALSE
        )
        
        # Tulis judul di sel A1 saja
        openxlsx::writeData(wb, sheet = 1, x = table_title, startRow = 1, startCol = 1, colNames = FALSE)
        openxlsx::addStyle(wb, sheet = 1, style = title_style, rows = 1, cols = 1)
        
        
        # Tulis data (mulai baris ke-2)
        openxlsx::writeData(wb, sheet = 1, x = data_to_download, startRow = 2, headerStyle = header_style)
        
        # Tambahkan border ke isi data
        openxlsx::addStyle(
          wb, sheet = 1, style = data_style,
          rows = 3:(nrow(data_to_download) + 2),
          cols = 1:ncol(data_to_download),
          gridExpand = TRUE,
          stack = TRUE
        )
        
        # ======== TAMBAHKAN CATATAN DI BAWAH TABEL ========
        # Ambil waktu saat ini
        timestamp <- format(Sys.time(), "%d %B %Y pukul %H:%M WIB", tz = "Asia/Jakarta")
        
        # Teks catatan
        footer_text <- paste("Sumber: BPS Kabupaten Karimun (data diolah), diakses pada", timestamp)
        
        # Tentukan baris tempat catatan ditulis (baris setelah data terakhir + 2)
        footer_row <- nrow(data_to_download) + 4  # 1 baris kosong setelah tabel
        
        # Tulis catatan di kolom pertama
        openxlsx::writeData(wb, sheet = 1, x = footer_text, startRow = footer_row, startCol = 1, colNames = FALSE)
        
        # Tambahkan style miring & kecil
        footer_style <- openxlsx::createStyle(
          fontSize = 9,
          textDecoration = "italic",
          halign = "left"
        )
        
        openxlsx::addStyle(wb, sheet = 1, style = footer_style, rows = footer_row, cols = 1, gridExpand = TRUE)
        
        
        # Atur lebar kolom
        get_width <- function(col) {
          max_len <- max(nchar(as.character(col)), na.rm = TRUE)
          return(min(max(10, max_len + 2), 30))
        }
        
        col_widths <- sapply(data_to_download, get_width)
        openxlsx::setColWidths(wb, sheet = 1, cols = 1:ncol(data_to_download), widths = col_widths)
        
        # Simpan
        openxlsx::saveWorkbook(wb, file, overwrite = TRUE)
      }
    )
    
    
    ## Download semua tabel
    output$downloadAllNeraca <- downloadHandler(
      filename = function() {
        "Semua Tabel Neraca Perdagangan.zip"
      },
      content = function(file) {
        # Buat folder sementara
        temp_dir <- tempdir()
        
        title_style <- openxlsx::createStyle(
          fontSize = 14,
          textDecoration = "bold",
          halign = "center"
        )
        
        # Daftar nama file dan tabel beserta judul tabel
        daftar_tabel <- list(
          "1. Nilai Neraca Perdagangan.xlsx"=list(tabel = tabel1_neraca, judul = "Tabel 1. Nilai Neraca Perdagangan"),
          "2. Perkembangan Nilai Neraca Perdagangan.xlsx"=list(tabel = tabel2_neraca, judul = "Tabel 2. Perkembangan Nilai Neraca Perdagangan"),
          "3. Perkembangan Nilai Neraca Perdagangan (c-t-c).xlsx"=list(tabel = tabel3_neraca, judul = "Tabel 3. Perkembangan Nilai Neraca Perdagangan")
        )
        
        clean_colnames <- function(names_vec) {
          names_vec <- gsub("\\s+", " ", names_vec)
          names_vec <- trimws(names_vec)
          names_vec <- gsub("(?<![a-zA-Z0-9])-\\s*-\\s*(?![a-zA-Z0-9])", " - ", names_vec, perl = TRUE)
          return(names_vec)
        }
        
        format_numeric <- function(df) {
          numeric_cols <- sapply(df, is.numeric)
          df[numeric_cols] <- lapply(df[numeric_cols], function(x) {
            out <- formatC(x, format = "f", digits = 2, big.mark = ".", decimal.mark = ",")
            out[is.na(x)] <- ""
            out
          })
          return(df)
        }
        
        
        
        write_table_to_excel <- function(data, file_path, judul_tabel) {
          colnames(data) <- clean_colnames(colnames(data))
          data <- format_numeric(data)
          
          wb <- openxlsx::createWorkbook()
          openxlsx::addWorksheet(wb, "Sheet1")
          
          header_style <- openxlsx::createStyle(
            textDecoration = "bold",
            wrapText = TRUE,
            halign = "center",
            valign = "center",
            border = "TopBottomLeftRight",
            borderStyle = "thin"
          )
          
          data_style <- openxlsx::createStyle(
            border = "TopBottomLeftRight",
            borderStyle = "thin"
          )
          
          judul_style <- openxlsx::createStyle(
            textDecoration = "bold",
            halign = "left",
            valign = "center",
            fontSize = 12,
            wrapText = FALSE
          )
          
          openxlsx::writeData(wb, sheet = 1, x = judul_tabel, startRow = 1, startCol = 1, colNames = FALSE)
          openxlsx::addStyle(wb, sheet = 1, style = judul_style, rows = 1, cols = 1)
          
          openxlsx::writeData(wb, sheet = 1, x = data, startRow = 2, headerStyle = header_style)
          openxlsx::addStyle(
            wb, sheet = 1, style = data_style,
            rows = 3:(nrow(data) + 2),
            cols = 1:ncol(data),
            gridExpand = TRUE,
            stack = TRUE
          )
          
          # ===== Tambahkan footer =====
          timestamp <- format(Sys.time(), "%d %B %Y pukul %H:%M WIB", tz = "Asia/Jakarta")
          footer_text <- paste("Sumber: BPS Kabupaten Karimun (data diolah), diakses pada", timestamp)
          footer_row <- nrow(data) + 4
          
          openxlsx::writeData(wb, sheet = 1, x = footer_text, startRow = footer_row, startCol = 1, colNames = FALSE)
          footer_style <- openxlsx::createStyle(
            fontSize = 9,
            textDecoration = "italic",
            halign = "left"
          )
          openxlsx::addStyle(wb, sheet = 1, style = footer_style, rows = footer_row, cols = 1)
          
          get_width <- function(col) {
            max_len <- max(nchar(as.character(col)), na.rm = TRUE)
            return(min(max(10, max_len + 2), 30))
          }
          col_widths <- sapply(data, get_width)
          openxlsx::setColWidths(wb, sheet = 1, cols = 1:ncol(data), widths = col_widths)
          
          openxlsx::saveWorkbook(wb, file_path, overwrite = TRUE)
        }
        
        file_paths <- c()
        for (nama_file in names(daftar_tabel)) {
          path_file <- file.path(temp_dir, nama_file)
          write_table_to_excel(daftar_tabel[[nama_file]]$tabel, path_file, daftar_tabel[[nama_file]]$judul)
          file_paths <- c(file_paths, path_file)
        }
        
        zip::zipr(zipfile = file, files = file_paths)
      },
      contentType = "application/zip"  # <--- INI PENTING!
    )
    
    
    
    output$tabel_data_neraca_ekspor <- renderDT({
      req(data_ekspor)  # Pastikan data tersedia
      
      data_sorted <- data_ekspor %>%
        arrange(desc(TAHUN), desc(BULAN)) %>%
        select(-NILAI_JUTA)
      
      datatable(
        data_sorted,
        options = list(
          scrollX = TRUE,
          scrollY = "300px",
          fixedHeader = TRUE,
          ordering = TRUE,
          columnDefs = list(
            list(targets = 0, className = "dt-nowrap")
          )
        ),
        rownames = FALSE
      )
    })
    
    
    
    output$tabel_data_neraca_impor <- renderDT({
      req(data_impor)  # Pastikan data tersedia
      
      data_sorted <- data_impor %>%
        arrange(desc(TAHUN), desc(BULAN))%>%
        select(-NILAI_JUTA)
      
      datatable(
        data_sorted,
        options = list(
          scrollX = TRUE,
          scrollY = "300px",
          fixedHeader = TRUE,
          ordering = TRUE,
          columnDefs = list(
            list(targets = 0, className = "dt-nowrap")
          )
        ),
        rownames = FALSE
      )
    })
    
    
    
    output$tabel_output_neraca_3 <- renderDT({
      
      data_terpilih <- switch(input$pilih_tabel_neraca,
                              "1. Nilai Neraca Perdagangan"=tabel1_neraca,
                              "2. Perkembangan Nilai Neraca Perdagangan"=tabel2_neraca,
                              "3. Perkembangan Nilai Neraca Perdagangan (c-t-c)"=tabel3_neraca
      )
      
      numerik_cols <- sapply(data_terpilih, is.numeric)
      data_terpilih[numerik_cols] <- lapply(
        data_terpilih[numerik_cols],
        function(x) {
          x_formatted <- formatC(x, format = "f", digits = 2, big.mark = ".", decimal.mark = ",")
          x_formatted[is.na(x)] <- ""
          x_formatted
        }
      )
      
      
      
      datatable(
        data_terpilih,
        options = list(
          scrollX = TRUE,
          scrollY = "300px",
          scrollCollapse = TRUE,  # Agar tinggi menyesuaikan jika baris sedikit
          fixedHeader = TRUE,
          ordering = FALSE,
          columnDefs = list(
            list(targets = 0, className = "dt-nowrap")  # Hindari wrap teks kolom pertama
          )
        ),
        rownames = FALSE
      )
      
      
    })
    
  })
  
  
  ##Tombol kirim saran
  sheet_id <- "1ZcRZ459NWXRO308ESegrwb2zjQZOMk47FZ07W7PIx88"
  
  
  # Autentikasi menggunakan token yang sudah disimpan

  # gs4_auth(token = readRDS("gs4_token.rds"))
  ##gs4_auth(path = "/srv/shiny-server/sipro/gs4-sa.json")
  # gs4_deauth()
  # gs4_user()



  count_words <- function(text) {
    if (is.null(text) || text == "") return(0)
    words <- unlist(strsplit(trimws(text), "\\s+"))
    words <- words[words != ""]
    length(words)
  }
  
  observe({
    nama_ok <- count_words(input$namaPengguna) >= 1
    saran_ok <- count_words(input$masukanPengguna) >= 5
    shinyjs::toggleState("kirimSaran", condition = nama_ok && saran_ok)
  })
  
  observeEvent(input$kirimSaran, {
    req(input$namaPengguna, input$masukanPengguna)
    
    new_data <- data.frame(
      Tanggal = format(Sys.time(), tz = "Asia/Jakarta", usetz = TRUE),
      Nama = input$namaPengguna,
      Masukan = input$masukanPengguna,
      stringsAsFactors = FALSE
    )
    
    
    tryCatch({
      # Pastikan autentikasi sudah jalan sebelum ini
      sheet_append(sheet_id, new_data)
      
      shinyalert("Berhasil!", "Terima kasih atas masukan Anda.", type = "success")
      
      updateTextInput(session, "namaPengguna", value = "")
      updateTextAreaInput(session, "masukanPengguna", value = "")
      
    }, error = function(e) {
      shinyalert("Gagal", paste("Error:", e$message), type = "error")
    })
  })
  
  
  
  
  # 1. UI Dinamis untuk input file
  output$file_inputs_ui <- renderUI({
    req(input$jumlah_file)
    lapply(1:input$jumlah_file, function(i) {
      fileInput(
        inputId = paste0("file_input_", i),
        label = paste("Pilih File", i, "(.sav)"),
        accept = ".sav"
      )
    })
  })
  
  # ReactiveValues untuk simpan data dan state modal
  data_storage <- reactiveValues(data_list = list())
  pending_index <- reactiveVal(NULL)
  pending_data <- reactiveVal(NULL)
  modal_active <- reactiveVal(FALSE)  # Flag modal sedang aktif
  
  # Observer untuk file input dinamis sesuai jumlah_file
  observeEvent(input$jumlah_file, {
    req(input$jumlah_file)
    
    lapply(1:input$jumlah_file, function(i) {
      observeEvent(input[[paste0("file_input_", i)]], {
        req(input[[paste0("file_input_", i)]])
        
        # Jangan proses kalau modal sedang aktif (supaya gak kedip modal)
        if (modal_active()) return()
        
        file <- input[[paste0("file_input_", i)]]
        data_sav <- haven::read_sav(file$datapath)
        
        # Kolom wajib
        required_cols <- c("BULAN", "NILAI", "BERAT", "PELABUHAN",
                           "HS", "NEGARA", "JENIS", "SEKTOR", "KOMODITI",
                           "HS2", "HSDUA", "NEGARA2", "PELABUHAN1")
        
        if (!all(required_cols %in% names(data_sav))) {
          # Tampilkan notifikasi error
          showNotification(
            paste0("Data yang diupload bukan merupakan data ekspor/impor"),
            type = "error",
            duration = 8
          )
          
          modal_active(TRUE)  # Optional: flag jika kamu gunakan untuk mengunci proses lanjutan
          return()            # Stop proses untuk file ini
        }
        
        
        if (!"TAHUN" %in% names(data_sav)) {
          pending_data(data_sav)
          pending_index(i)
          
          modal_active(TRUE)  # Set dulu sebelum show modal
          showModal(modalDialog(
            title = paste("File", i, ": Kolom TAHUN tidak ditemukan"),
            numericInput(
              "tahun_input_modal",
              "Masukkan Tahun (4 digit):",
              value = NA,
              min = 2024,
              width = "100%"
            ),
            footer = tagList(
              actionButton(
                inputId = "confirm_tahun_button",
                label = "Konfirmasi",
                style = "width: 100%;",
                class = "btn btn-primary"
              )
              
            ),
            easyClose = FALSE
          ))
        } else {
          data_storage$data_list[[i]] <- data_sav
        }
      }, ignoreInit = TRUE)
    })
  })
  
  # Event konfirmasi input tahun dari modal
  observeEvent(input$confirm_tahun_button, {
    req(pending_data(), input$tahun_input_modal)
    
    tahun_val <- input$tahun_input_modal
    valid <- !is.na(tahun_val) &&
      nchar(as.character(tahun_val)) == 4 &&
      tahun_val >= 2024
    
    if (valid) {
      i <- pending_index()
      data_sav <- pending_data()
      
      data_sav$TAHUN <- as.integer(tahun_val)
      data_storage$data_list[[i]] <- data_sav
      
      # Reset pending data & index
      pending_data(NULL)
      pending_index(NULL)
      
      removeModal()
      modal_active(FALSE)  # Setelah remove modal
      
    } else {
      showNotification("Masukkan tahun 4 digit yang valid", type = "error")
    }
  })
  
  
  
  
  
  # 5. ReactiveVal untuk data bersih gabungan
  data_bersih <- reactiveVal(NULL)
  
  # 6. Proses cleaning data saat tombol ditekan
  observeEvent(input$clean_button, {
    req(length(data_storage$data_list) == input$jumlah_file)
    
    # Proses pembersihan per file
    data_bersih_list <- lapply(data_storage$data_list, function(df) {
      df$BULAN <- as.numeric(df$BULAN)
      df$TAHUN <- as.numeric(df$TAHUN)
      
      df$PROPINSI <- trimws(as.character(df$PROPINSI), which = "both")
      df$PELABUHAN <- trimws(as.character(df$PELABUHAN), which = "both")
      
      df$HS <- trimws(as.character(df$HS), which = "both")
      df$HS <- ifelse(nchar(df$HS) == 7, paste0("0", df$HS), df$HS)
      
      df$NEGARA <- trimws(as.character(df$NEGARA), which = "both")
      
      df$BERAT <- as.numeric(df$BERAT)
      df$NILAI <- as.numeric(df$NILAI)
      
      df$HSDUA <- trimws(as.character(df$HSDUA), which = "both")
      df$HSDUA <- ifelse(nchar(df$HSDUA) == 1, paste0("0", df$HSDUA), df$HSDUA)
      
      df$KABKOT <- trimws(as.character(df$KABKOT), which = "both")
      df$PELABUHAN1 <- trimws(as.character(df$PELABUHAN1), which = "both")
      df$NEGARA2 <- trimws(as.character(df$NEGARA2), which = "both")
      df$HS2 <- trimws(as.character(df$HS2), which = "both")
      df$JENIS <- trimws(as.character(df$JENIS), which = "both")
      df$SEKTOR <- trimws(as.character(df$SEKTOR), which = "both")
      df$KOMODITI <- trimws(as.character(df$KOMODITI), which = "both")
      
      return(df)
    })
    
    # Gabungkan semua data yang sudah dibersihkan
    gabungan <- dplyr::bind_rows(data_bersih_list)
    gabungan <- dplyr::arrange(gabungan, desc(TAHUN), desc(BULAN))
    
    data_bersih(gabungan)
    
    shinyjs::disable("downloadDataClean")
    
    output$downloadDataClean <- downloadHandler(
      filename = function() {
        paste0("Data Bersih", ".sav")
      },
      content = function(file) {
        req(data_bersih())
        haven::write_sav(data_bersih(), file)
      }
    )
    
    observe({
      if (!is.null(data_bersih()) && nrow(data_bersih()) > 0) {
        shinyjs::enable("downloadDataClean")
      } else {
        shinyjs::disable("downloadDataClean")
      }
    })
  })
  
  
  # 7. Render tabel hasil data bersih
  output$tabel_output_clean_3 <- DT::renderDT({
    req(data_bersih())
    
    DT::datatable(
      data_bersih(),
      options = list(
        scrollX = TRUE,
        scrollY = "300px",
        scrollCollapse = TRUE,
        fixedHeader = TRUE,
        ordering = TRUE,
        columnDefs = list(
          list(targets = 0, className = "dt-nowrap")
        )
      ),
      rownames = FALSE
    )
  })
  
  
  # Reactive value untuk menyimpan status kode salah
  kode_salah_paparan <- reactiveVal(FALSE)
  
  # Tampilkan modal saat tombol ditekan
  observeEvent(input$akses_paparan_button, {
    kode_salah_paparan(FALSE)  # Reset status
    
    showModal(modalDialog(
      title = div(
        style = "text-align: center;",
        tags$h4("Kode Akses Diperlukan", style = "margin: 0;")
      ),
      
      passwordInput(
        inputId = "kode_akses",
        label = "Silakan masukkan kode akses untuk melihat dan upload publikasi:",
        width = "100%"
      ),
      
      uiOutput("pesan_kode_salah_paparan"),
      
      footer = div(
        style = "display: flex; justify-content: space-between; gap: 10px; flex-wrap: wrap;",
        actionButton("batal_kode_paparan", "Batal", class = "btn btn-outline-secondary", style = "flex: 1; min-width: 120px;"),
        actionButton(
          "konfirmasi_kode", 
          "Kirim", 
          class = "btn btn-primary", 
          style = "flex: 1; min-width: 120px;", 
          onclick = "
    Shiny.setInputValue(
      'kode_submit', 
      {
        kode: document.getElementById('kode_akses').value,
        nonce: Math.random()
      }, 
      { priority: 'event' }
    );
  "
        )
        
        
      ),
      
      easyClose = FALSE,
      
      # Tambahkan script untuk fokus dan enter
      tags$script(HTML("
      $(document).on('shown.bs.modal', function() {
        var input = document.getElementById('kode_akses');
        if (input && !input.dataset.listenerAttached) {
          input.focus();
          input.addEventListener('keypress', function(e) {
            if (e.key === 'Enter') {
              document.getElementById('konfirmasi_kode').click();
            }
          });
          input.dataset.listenerAttached = 'true';
        }
      });
    "))
    ))
  })
  
  
  # Tindakan jika tombol Batal ditekan
  observeEvent(input$batal_kode_paparan, {
    removeModal()
  })
  
  # Pesan kesalahan ditampilkan langsung di modal
  output$pesan_kode_salah_paparan <- renderUI({
    if (kode_salah_paparan()) {
      div(style = "color: red; margin-top: 10px;", "Kode akses yang Anda masukkan tidak sesuai.")
    }
  })
  
  # Validasi kode akses
  observeEvent(input$kode_submit, {
    req(input$kode_submit$kode)
    
    if (input$kode_submit$kode == "karimunjaya2101") {
      kode_salah_paparan(FALSE)
      removeModal()
      runjs("window.open('https://drive.google.com/drive/folders/1RIJYqI2st8o7I-gntG2713Ask5iEtI0y?usp=sharing', '_blank');")
    } else {
      kode_salah_paparan(TRUE)
    }
  })
  
  
  
  
  
  # Reactive value untuk status kode salah
  kode_salah_data_historis <- reactiveVal(FALSE)
  
  # Tampilkan modal saat infoBox (actionLink) diklik
  observeEvent(input$akses_data_historis, {
    kode_salah_data_historis(FALSE)  # Reset status
    
    showModal(modalDialog(
      title = div(
        style = "text-align: center;",
        tags$h4("Kode Akses Diperlukan", style = "margin: 0;")
      ),
      
      passwordInput(
        inputId = "kode_akses_data_historis",
        label = "Silakan masukkan kode akses untuk melihat kumpulan data historis:",
        width = "100%"
      ),
      
      uiOutput("pesan_kode_salah_data_historis"),
      
      footer = div(
        style = "display: flex; justify-content: space-between; gap: 10px; flex-wrap: wrap;",
        actionButton("batal_kode_data_historis", "Batal", class = "btn btn-outline-secondary", style = "flex: 1; min-width: 120px;"),
        actionButton(
          "konfirmasi_kode_data_historis", 
          "Kirim", 
          class = "btn btn-primary", 
          style = "flex: 1; min-width: 120px;",
          onclick = "
    Shiny.setInputValue(
      'kode_submit_data_historis',
      {
        kode: document.getElementById('kode_akses_data_historis').value,
        nonce: Math.random()
      },
      { priority: 'event' }
    );
  "
        )
        
        
      ),
      
      easyClose = FALSE,
      
      # Script yang benar-benar menunggu modal selesai tampil
      tags$script(HTML("
      $(document).on('shown.bs.modal', function() {
        var input = document.getElementById('kode_akses_data_historis');
        if (input) {
          input.focus();
          input.addEventListener('keypress', function(e) {
            if (e.key === 'Enter') {
              document.getElementById('konfirmasi_kode_data_historis').click();
            }
          });
        }
      });
    "))
    ))
  })
  
  
  
  # Tindakan jika tombol Batal ditekan
  observeEvent(input$batal_kode_data_historis, {
    removeModal()
  })
  
  # Render pesan jika salah
  output$pesan_kode_salah_data_historis <- renderUI({
    if (kode_salah_data_historis()) {
      div(style = "color: red; margin-top: 10px;", "Kode akses yang Anda masukkan tidak sesuai.")
    }
  })
  
  
  
  observeEvent(input$kode_submit_data_historis, {
    req(input$kode_submit_data_historis$kode)
    
    if (input$kode_submit_data_historis$kode == "karimunjaya2101") {
      kode_salah_data_historis(FALSE)
      removeModal()
      runjs("window.open('https://1drv.ms/f/c/9e7e9ee182eb06e5/EmCtERpJiqtBj4dKBRhmdRUB8GEAHBRXYokXrX0moxxVKQ?e=EfyV1Z', '_blank');")
    } else {
      kode_salah_data_historis(TRUE)
    }
  })
  
  
  
  
  
  # Reactive value untuk status kode salah
  kode_salah_saran <- reactiveVal(FALSE)
  
  # Tampilkan modal saat tombol diklik
  observeEvent(input$lihatSaran, {
    kode_salah_saran(FALSE)  # Reset status
    
    showModal(modalDialog(
      title = div(
        style = "text-align: center;",
        tags$h4("Kode Akses Diperlukan", style = "margin: 0;")
      ),
      
      passwordInput(
        inputId = "kodeAkses",
        label = "Silakan masukkan kode akses untuk melihat saran pengguna:",
        width = "100%"
      ),
      
      uiOutput("pesan_kode_salah_saran"),
      
      footer = div(
        style = "display: flex; justify-content: space-between; gap: 10px; flex-wrap: wrap;",
        actionButton("batalKodeSaran", "Batal", class = "btn btn-outline-secondary", style = "flex: 1; min-width: 120px;"),
        actionButton(
          "konfirmasiKode", 
          "Kirim", 
          class = "btn btn-primary", 
          style = "flex: 1; min-width: 120px;",
          onclick = "
    Shiny.setInputValue(
      'kode_submit_saran', 
      {
        kode: document.getElementById('kodeAkses').value,
        nonce: Math.random()
      },
      { priority: 'event' }
    );
  "
        )
        
        
      ),
      
      easyClose = FALSE,
      
      # Tambahan JavaScript agar fokus langsung dan bisa tekan Enter
      tags$script(HTML("
      $(document).on('shown.bs.modal', function() {
        var input = document.getElementById('kodeAkses');
        if (input && !input.dataset.listenerAttached) {
          input.focus();
          input.addEventListener('keypress', function(e) {
            if (e.key === 'Enter') {
              document.getElementById('konfirmasiKode').click();
            }
          });
          input.dataset.listenerAttached = 'true';
        }
      });
    "))
    ))
  })
  
  
  # Tindakan jika tombol Batal ditekan
  observeEvent(input$batalKodeSaran, {
    removeModal()
  })
  
  # Render pesan kesalahan
  output$pesan_kode_salah_saran <- renderUI({
    if (kode_salah_saran()) {
      div(style = "color: red; margin-top: 10px;", "Kode akses yang Anda masukkan tidak sesuai.")
    }
  })
  
  # Validasi kode akses
  observeEvent(input$kode_submit_saran, {
    req(input$kode_submit_saran$kode)
    
    if (input$kode_submit_saran$kode == "karimunjaya2101") {
      kode_salah_saran(FALSE)
      removeModal()
      runjs("window.open('https://docs.google.com/spreadsheets/d/1ZcRZ459NWXRO308ESegrwb2zjQZOMk47FZ07W7PIx88/edit?usp=sharing', '_blank');")
    } else {
      kode_salah_saran(TRUE)
    }
  })
  
  
  
}

shinyApp(ui, server)

