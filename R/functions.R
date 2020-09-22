#' Perform all checks
check_sheets <- function(inf, dat){
  cn_inf <- colnames(inf)
  cn_dat <- colnames(dat)

  # Check colnames
  if(!identical(cn_inf, c("Empreendimento", "Código do cliente",
                          "Relatório", "Revisão",
                          "responsável", "Unidade avaliada",
                          "Estação", "Cidade de avaliação", "zona bioclimática",
                          "temperatura do dia típico"))
  ){stop("Existe algum problema com a planilha! Utilize o template disponível no link:
            https://docs.google.com/spreadsheets/d/1OGM4r6rmo3G8L49aPlLnHGM7Nwy7OJMcmduleYbUp20/edit?usp=sharing")}

  if(cn_dat[1] != "cenario" | cn_dat[2] != "Hora"){
    stop("A planilha dat tem alguma problema nas colunas: cenario ou Hora")
  }

  # Check for missing values
  if(length(inf) != 10 | any(is.na(inf))){
    stop("Existe um problema na planilha `info`. Verifique no google drive! Provavelmente uma céulua foi deixada em branco!")
  }
}


#' Load and check data
load_save_data <- function(link, root) {

  if(!dir.exists(paths = root)){dir.create(path = root)}

  # Checks---------------------------
  inf <- googlesheets4::read_sheet(ss = link, sheet = 1)
  dat <- googlesheets4::read_sheet(ss = link, sheet = 2)

  check_sheets(inf, dat)

  # Create folder
  folder <- paste(inf$`Código do cliente`, sep = "/")
  folder_2 <- paste(inf$Relatório)
  subfolder <-
    paste(paste(paste("R0", inf$Revisão, sep = ""),
                #substr(inf$Estação, start = 0, stop = 3),
                abbreviate(inf$`Unidade avaliada`),
                Sys.Date(),  sep = "_"))
  file_name <-
    paste(substr(inf$`Código do cliente`, start = 0, stop = 6),
          #substr(inf$Relatório, start = 0 , stop = 3),
          gsub(x = inf$Relatório, pattern = " ", replacement = "_"),
          paste("R0", inf$Revisão, sep = ""),
          substr(inf$Estação, start = 0, stop = 3),
          sep = "_")

  # Create folder with client id number
  if(!dir.exists(paths = paste(root, folder, sep = "/"))){
    dir.create(path = paste(root, folder, sep = "/"))
  }
  if(!dir.exists(paths = paste(root, folder, folder_2, sep = "/"))){
    dir.create(path = paste(root, folder, folder_2, sep = "/"))
  }

  file_path <- paste(root, folder, folder_2, sep = "/")

  # Save raw data
  raw_data <- list(inf = inf, dat = dat, folder = folder,
                   file_name = file_name,
                   fine_path = file_path)

  # Save XLSX copy
  writexl::write_xlsx(x = raw_data[1:2],
             path = paste(file_path, "/",
                          file_name, "_dados_brutos.xlsx", sep = ""))
  return(raw_data)
}

#' Relatório NBR
#'
#' @param link "The google sheet sharing link"
#' @param root "The name of the root folder to store all output
#'
#' @return Several plots and tables as a result of the NBR norm.
#' @export
relatorio <- function(link, root = "Rel_NBR"){
  raw <- load_save_data(link = link, root)
  inf <- raw$inf
  dat <- raw$dat
  tdt  <- inf$`temperatura do dia típico`

  # Definir valores de referência
  # Verão
  if (inf$Estação == "Verão" & inf$`zona bioclimática` %in% c(1:7)){
    # Valores temp referência
    ref_min <- tdt
    ref_int <- tdt - 2
    ref_sup <- tdt - 4
  }
  if (inf$Estação == "Verão" & inf$`zona bioclimática` %in% c(8)){
    # Valores temp referência
    ref_min <- tdt
    ref_int <- tdt - 1
    ref_sup <- tdt - 2
  }
  if (inf$Estação == "Inverno" & inf$`zona bioclimática` %in% c(1:5)){
    # Valores temp referência
    ref_min <- tdt + 3
    ref_int <- tdt + 5
    ref_sup <- tdt + 7
  }

  folder <- raw$folder
  file_name <- raw$file_name
  file_path <- raw$fine_path

  est <- inf$Estação

  # 1. transpose data----------------------------------------------------------------
  dat$Hora <- lubridate::hour(dat$Hora)
  dat_long <- tidyr::pivot_longer(dat, -c("Hora", "cenario"))
  dat_long$value <- round(dat_long$value, digits = 2)

  # 2. Parameters--------------------------------------------------------------------
  # Valores temp referência
  #ref_min <- ref$temp[ref$ref == "mínimo"]
  #ref_int <- ref$temp[ref$ref == "intermediário"]
  #ref_sup <- ref$temp[ref$ref == "superior"]

  temp_range <- range(dat_long$value)

  if(est == "Inverno"){
    temp_max   <- ceiling(max(dat_long$value)) + 0.5
    temp_min   <- floor(min(dat_long$value)) - 2
    if(temp_min > ref_min) {
      temp_min <- ref_min - 1
    }
  }
  if(est == "Verão"){
  temp_max     <- ceiling(max(dat_long$value) + 2)
  temp_min     <- floor(min(dat_long$value))
  if(temp_max < ref_min) {
    temp_max <- ref_min + 1
  }
  }

  ambiente_ord <- colnames(dat)[-c(1,2)]
  n_ambient    <- length(ambiente_ord)
  n_cenario    <- length(unique(dat$cenario))

  # Ref table
  if(est == "Inverno"){
    ref_tab <-
      data.frame(
        classe = c("Não atende", "Mínimo", "Intermediário", "Superior"),
        tstart = c(temp_min, ref_min, ref_int, ref_sup),
        tend   = c(ref_min, ref_int, ref_sup, temp_max),
        colors = c("#ffc68f", "#ffffd1", "#c0e7bd", "#53a9ce"))
  }
  if(est == "Verão"){
    ref_tab <-
      data.frame(
        classe = c("Não atende", "Mínimo", "Intermediário", "Superior"),
        tstart = c(ref_min, ref_int, ref_sup, temp_min),
        tend   = c(temp_max, ref_min, ref_int, ref_sup),
        colors = c("#ffc68f", "#ffffd1", "#c0e7bd", "#53a9ce"))
  }


  ref_tab$colors <- factor(x = ref_tab$colors,
                           levels = c("#ffc68f", "#ffffd1", "#c0e7bd", "#53a9ce"))
  ref_tab$classe <- factor(x = ref_tab$classe,
                           levels = c("Não atende", "Mínimo", "Intermediário", "Superior"))

  # Final result-------------------------------------------------------------------
  if (est == "Inverno"){
    dat_long  %>% group_by(name, cenario) %>%
      summarise(res = min(value)) %>%
      mutate(classe = "") -> res_tab
  }
  if (est == "Verão"){
    dat_long  %>% group_by(name, cenario) %>%
      summarise(res = max(value)) %>%
      mutate(classe = "") -> res_tab
  }

  # Classe por ambiente
  if(est == "Inverno"){
    res_tab[res_tab$res < ref_min, ]$classe <- "Não atende"
    res_tab[res_tab$res >= ref_min & res_tab$res < ref_int, ]$classe <- "Mínimo"
    res_tab[res_tab$res >= ref_int & res_tab$res < ref_sup, ]$classe <- "Intermediário"
    res_tab[res_tab$res >= ref_sup, ]$classe <- "Superior"
  }


  if(est == "Verão") {
    res_tab[res_tab$res > ref_min, ]$classe <- "Não atende"
    res_tab[res_tab$res <= ref_min & res_tab$res > ref_int, ]$classe <- "Mínimo"
    res_tab[res_tab$res <=ref_int & res_tab$res > ref_sup, ]$classe <- "Intermediário"
    res_tab[res_tab$res <= ref_sup, ]$classe <- "Superior"
  }
  res_tab %>% select(Ambiente = name,
                     "Cenário" = cenario,
                     "Temp Max" = res,
                     "Desempenho" = classe) -> tabela_final
  # Tabela com nomes
  res_tab %>%
    select(-res) %>%
    tidyr::pivot_wider(names_from = cenario, values_from = classe) %>%
    select(Ambiente = name, everything()) %>%
    ungroup() -> tabela_final_names

  res_tab %>%
    select(-classe) %>%
    tidyr::pivot_wider(names_from = cenario, values_from = res) %>%
    select(Ambiente = name, everything()) %>%
    ungroup() -> tabela_final_number

  # Tabela para exportar
  ftc  <- flextable::flextable(tabela_final_names)  # With classes
  ftn  <- flextable::flextable(tabela_final_number) # with numbers
  cn        <- colnames(tabela_final_names)

  # Color by class
  tab_num <- tabela_final_number %>% select(-Ambiente)
  #tab_nam <- tabela_final_names %>% select(-Ambiente)

  if (est == "Inverno"){
    c_fail <- tab_num < ref_min
    c_mini <- tab_num >= ref_min & tab_num < ref_int
    c_inti <- tab_num >= ref_int & tab_num < ref_sup
    c_supe <- tab_num >= ref_sup
  }
  if (est == "Verão"){
    c_fail <- tab_num > ref_min
    c_mini <- tab_num <= ref_min & tab_num > ref_int
    c_inti <- tab_num <= ref_int & tab_num > ref_sup
    c_supe <- tab_num <= ref_sup
  }

  # Color table with names
  for(jj in 1:ncol(tab_num)){
    ftc <- flextable::bg(ftc, i = c_fail[, jj], j = jj + 1, bg = as.character(ref_tab$colors[1]))
    ftc <- flextable::bg(ftc, i = c_mini[, jj], j = jj + 1, bg = as.character(ref_tab$colors[2]))
    ftc <- flextable::bg(ftc, i = c_inti[, jj], j = jj + 1, bg = as.character(ref_tab$colors[3]))
    ftc <- flextable::bg(ftc, i = c_supe[, jj], j = jj + 1, bg = as.character(ref_tab$colors[4]))
  }
  # align
  ftc <-  flextable::align(ftc, align =  "center", part = "all")
  ftc <-  flextable::width(ftc, j = 2:n_cenario, width = 1.5)

  # Color table with numbers
  for(jj in 1:ncol(tab_num)){
    ftn <- flextable::bg(ftn, i = c_fail[, jj], j = jj + 1, bg = as.character(ref_tab$colors[1]))
    ftn <- flextable::bg(ftn, i = c_mini[, jj], j = jj + 1, bg = as.character(ref_tab$colors[2]))
    ftn <- flextable::bg(ftn, i = c_inti[, jj], j = jj + 1, bg = as.character(ref_tab$colors[3]))
    ftn <- flextable::bg(ftn, i = c_supe[, jj], j = jj + 1, bg = as.character(ref_tab$colors[4]))
  }
  # align
  ftn <- flextable::align(ftn, align =  "center", part = "all")
  ftn <- flextable::width(ftn, j = 2:n_cenario, width = 1.5)

  # Exportar tabela Final---------------------------------------------------------
  ftc %>%
    flextable::fontsize(size = 16, part = "header") %>%
    flextable::font(fontname = "Arial", part = "all") %>%
    flextable::fontsize(size = 14) -> final_tab_classe

  ftn %>%
    flextable::fontsize(size = 16, part = "header") %>%
    flextable::font(fontname = "Arial", part = "all") %>%
    flextable::fontsize(size = 14) -> final_tab_number

  # Salvar tabelas----------------------------------------------------------------
  openxlsx::write.xlsx(x  = tabela_final,
                       file = paste(file_path,
                                    paste(file_name, "tab_final",
                                          paste(inf$responsável, ".xlsx", sep = ""), sep = "_"),
                                    sep = "/"))
  flextable::save_as_image(x = final_tab_classe,
                paste(file_path,
                      paste(file_name, "tab_resumo",
                            paste(inf$responsável, ".png", sep = ""), sep = "_"),
                      sep = "/"))
  flextable::save_as_image(x = final_tab_number,
                paste(file_path,
                      paste(file_name, "tab_detalhe",
                            paste(inf$responsável, ".png", sep = ""), sep = "_"),
                      sep = "/"))
  # Plots-------------------------------------------------------------------------
  # Plot por hora----------------------------------------------------------------
  g_hora <-
    ggplot(data = dat_long) +
    geom_rect(data = ref_tab,
              aes(ymin = tstart, ymax = tend, fill = colors),
              xmin = -Inf, xmax = Inf, alpha = 1) +
    geom_hline(yintercept = ref_min, color = "tomato", size = 1.5) +

    scale_fill_identity(guide = guide_legend(title = ""),
                        labels = c(as.character(ref_tab$classe))) +
    geom_line(data = dat_long, aes(x= Hora, y = value, lty = cenario)) +
       #geom_text(data = dat_long %>% group_by(name) %>%
    #            filter(value == max(value)) %>%
    #            filter(Hora == max(Hora)),
    #          aes(y = value, x = Hora, label = value),
    #          nudge_x = -1.5, nudge_y = .5) +
    scale_linetype(name = "") +
    scale_x_continuous(limits = c(0, 24), breaks = seq(0,24,4)) +
    scale_y_continuous(limits = c(temp_min, temp_max),
                       breaks = seq(temp_min, temp_max, 1)) +

    facet_wrap(~name) +
    theme_classic(base_size = 14) +
    labs(y = "Temperatura (°C)",
         x = "Hora do dia",
         title = paste("Desempenho de", inf$Estação),
         subtitle = paste(inf$Empreendimento,
                          inf$`Unidade avaliada`, sep = " | "),
         caption = "@qualiabr") +
    theme(plot.title = element_text(size = 18),
          plot.subtitle = element_text(size = 14),
          strip.background = element_rect(size = .1),
          legend.position = "right")

  if (est == "Inverno") {
    g_hora <-
      g_hora +
      geom_point(data = dat_long %>% group_by(name) %>%
                 filter(value == min(value)) %>%
                 filter(Hora == min(Hora)),
               aes(y = value, x = Hora))
  }

  if (est == "Verão") {
    g_hora <-
      g_hora +
      geom_point(data = dat_long %>% group_by(name) %>%
                   filter(value == max(value)) %>%
                   filter(Hora == min(Hora)),
                 aes(y = value, x = Hora))
  }
  # Resultado Final----------------------------------------------------------
  res_tab %>%
    ungroup() %>%
    mutate(resultado = ifelse(classe == "Não atende", yes = "reprovado", no = "aprovado")) ->
    res_tab

  g_final <-
    ggplot(data = res_tab) +
    geom_rect(data = ref_tab,
              aes(ymin = tstart, ymax = tend, fill = colors),
              xmin = -Inf, xmax = Inf, alpha = 1) +
    geom_hline(yintercept = ref_min, color = "tomato", size = 1.5) +
    geom_segment(aes(x = name, y = res,
                     xend = name, yend = temp_min), color = gray(.3))  +
    scale_color_manual(values = c(gray(.2), "red")) +
    geom_text(aes(x = name, y = res, label = res),
              nudge_x = -.35, size = 3) +
    geom_point(aes(y = res, x = name),
               show.legend = FALSE, size = 3, alpha = 1, fill = gray(.2), shape = 21) +
    geom_point(aes(y = res, x = name, color = resultado),
               show.legend = FALSE, size = .7, alpha = 1) +

    scale_y_continuous(limits = c(temp_min, temp_max),
                       breaks = seq(temp_min, temp_max, 1)) +
    scale_x_discrete(limit = ambiente_ord[length(ambiente_ord):1]) +
    scale_fill_identity(guide = guide_legend(title = ""),
                        labels = as.character(ref_tab$classe)) +
    coord_flip() +
    facet_wrap(~ cenario) +
    theme_classic(base_size = 13) +
    labs(y = "Temperatura (°C)",
         x = "Ambiente",
         title = paste("Desempenho de", inf$Estação),
         subtitle = paste(inf$Empreendimento,
                          inf$`Unidade avaliada`, sep = " | "),
         caption = "@qualiabr") +
    theme(plot.title = element_text(size = 18),
          plot.subtitle = element_text(size = 14),
          strip.background = element_rect(size = .1),
          legend.position = "bottom",
          legend.key = element_rect(color = gray(.3), size = .5))

  # Save plots--------------------------------------------------------------------
  # plot dimentions
  width_final = 8
  height_final = 6
  width_hora = 10
  height_hora = 7.5

  # Números de cenários / i.g inverno
  if(n_cenario == 1){
    height_final = 4.5
    width_final = 5
    g_final <- g_final + theme(
      legend.text = element_text(size = 8)
    )
  }

  # Número de ambientes = 3
  if(n_ambient == 3){
    width_hora = 8
    height_hora = 4.5
  }
  # Número de ambientes = 4
  if(n_ambient == 4){
    width_hora = 8
    height_hora = 7.5
  }

  ggsave(plot = g_final,
         filename = paste(file_path,
                          paste(file_name, "fig_final",
                                paste(inf$responsável, ".png",sep = ""),
                                sep = "_"),
                          sep = "/"),
         height = height_final, width = width_final)

  ggsave(plot = g_hora,
         filename = paste(file_path,
                          paste(file_name, "fig_hora",
                                paste(inf$responsável, ".png",sep = ""),
                                sep = "_"),
                          sep = "/"),
         height = height_hora, width = width_hora)
}
