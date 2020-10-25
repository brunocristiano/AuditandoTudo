#Scrap ANEEL
setwd('E:/Dropbox/Bruno/Academico/UFMG - Estatística/9- Monografia/Estatística Aplicada a Auditoria/Dados')
#install.packages("RSelenium")
#install.packages("KeyboardSimulator")
library("KeyboardSimulator")
library("RSelenium")
library(XML)

#Código
{ 
  rD <- rsDriver(verbose = FALSE,port=4000L,chromever = "86.0.4240.22") #Abre a porta com o Chrome. Verifique qual sua versão do chrome
  remDr <- rD$client
  Sys.sleep(2)
  
  remDr$navigate("http://relatorios.aneel.gov.br/_layouts/xlviewer.aspx?id=/RelatoriosSAS/RelSAMPClasseConsNivel.xlsx&Source=http%3A%2F%2Frelatorios%2Eaneel%2Egov%2Ebr%2FRelatoriosSAS%2FForms%2FAllItems%2Easpx&DefaultItemOpen=1") 
  print("Aguarda carregar a página")
  remDr$maxWindowSize()
  
  remDr$setTimeout(type = "page load", milliseconds = 10000) 
  
  #Abre Ano
  print("Abre ano")
  webElem <- remDr$findElement(using = 'xpath', value = "//*[@id='afi.0.1']")
  webElem$highlightElement()                               
  webElem$clickElement()
  Sys.sleep(2)
  
  #Escolhe ano
  print("Escolhe ano")
  #Desmarca (Selecionar Tudo)
  webElem <- remDr$findElement(using = 'xpath', value = "//*[@id='ewaDialogBody']/div[2]/ul/li[1]/label")
  
  webElem$highlightElement()                               
  webElem$clickElement()
  Sys.sleep(1)
  
  #Seleciona 2020
  #2020 
  webElem <- remDr$findElement(using = 'xpath', value = "//*[@id='ewaDialogBody']/div[2]/ul/li[19]/label")
  #2019 webElem <- remDr$findElement(using = 'xpath', value = "//*[@id='ewaDialogBody']/div[2]/ul/li[18]/label")
  #2018 webElem <- remDr$findElement(using = 'xpath', value = "//*[@id='ewaDialogBody']/div[2]/ul/li[17]/label")
  webElem$highlightElement()                               
  webElem$clickElement()
  Sys.sleep(1)
  
  #Escolhe ok para ano
  print("Ok para ano")
  webElem <- remDr$findElement(using = 'xpath', value = "//*[@id='buttonarea']/button[1]")
  webElem$highlightElement()                               
  webElem$clickElement()
  Sys.sleep(5)
  
  
  for (i in 2:13) {
    
    #Abre o mês
    print("Abre mes")
    webElem <- remDr$findElement(using = 'xpath', value = "//*[@id='afi.1.1']")
    #webElem$highlightElement()                               
    webElem$clickElement()
    Sys.sleep(2)
    
    print("foi mes")
    if (i == 2) {
      #Se for no inicio
      #Escolhe (Selecionar Tudo) para desmarcar
      webElem <- remDr$findElement(using = 'xpath', value = "//*[@id='ewaDialogBody']/div[2]/ul/li[1]/label")
      #webElem$highlightElement()                               
      webElem$clickElement()
      Sys.sleep(1)
    }
    
    #Escolhe mês i
    Sys.sleep(2)
    print(paste0("fazendo mês ",i-1))
    
    if (i >= 3) {
      #Se nao for no inicio, tem que desmarcar o anterior
      print("Desmarca anterior")
      
      webElem <- remDr$findElement(using = 'xpath', value = paste0("//*[@id='ewaDialogBody']/div[2]/ul/li[",i-1,"]/label"))
      #webElem$highlightElement()                               
      webElem$clickElement() #Desmarca anterior
      Sys.sleep(1)
    }
    
    print("Marca o atual")
    webElem <- remDr$findElement(using = 'xpath', value = paste0("//*[@id='ewaDialogBody']/div[2]/ul/li[",i,"]/label"))
    #webElem$highlightElement()                               
    webElem$clickElement()
    Sys.sleep(1)
    
    #aperta Ok
    print("Aperta Ok")
    webElem <- remDr$findElement(using = 'xpath', value = "//*[@id='buttonarea']/button[1]")
    #webElem$highlightElement()                               
    webElem$clickElement()
    Sys.sleep(2)
    
    #Deseleciona Tudo
    
    for (k in 1:12) {
      
      #Filtra cada tipo de Classe
      #clica no Classe
      print("Clica no ClasseConsumo")
      webElem <- remDr$findElement(using = 'xpath', value = "//*[@id='afi.2.1']")
      #webElem$highlightElement()                               
      webElem$clickElement()
      Sys.sleep(1)
      
      
      if (k == 1) {
        #Se for no inicio
        #Escolhe (Selecionar Tudo) para desmarcar
        webElem <- remDr$findElement(using = 'xpath', value = "//*[@id='ewaDialogBody']/div[2]/ul/li[1]/label")
        #webElem$highlightElement()                               
        webElem$clickElement()
        Sys.sleep(1)
      }
      
      if (k == 1 && i >= 3) {
        #Se for no inicio
        #Escolhe (Selecionar Tudo) para desmarcar
        webElem <- remDr$findElement(using = 'xpath', value = "//*[@id='ewaDialogBody']/div[2]/ul/li[1]/label")
        #webElem$highlightElement()                               
        webElem$clickElement()
        Sys.sleep(1)
      }
      
      
      #Clica em cada classe com loop
      
      if (k >= 3) { #Se ele estiver no segundo item, tem que desmarcar o anterior, para depois marcar o segunte
        #Desmarcar o anterior
        print("Demarca o anterior")
        webElem <- remDr$findElement(using = 'xpath', value = paste0("//*[@id='ewaDialogBody']/div[2]/ul/li[",k-1,"]/label"))
        #webElem$highlightElement()                               
        webElem$clickElement()
        Sys.sleep(1)
      }
      
      if (k >= 2) {
        #Marca o atual
        print("Marca o atual")
        webElem <- remDr$findElement(using = 'xpath', value = paste0("//*[@id='ewaDialogBody']/div[2]/ul/li[",k,"]/label"))
        #webElem$highlightElement()                               
        webElem$clickElement()
        Sys.sleep(1)  
        
        #Da o ok!
        print("OK Classe")
        webElem <- remDr$findElement(using = 'xpath', value = "//*[@id='buttonarea']/button[1]")
        webElem$clickElement()
        Sys.sleep(1)  
        
        ################ Salva dados1 em cima
        print("Salva")
        webElem  <- remDr$findElement(using = "xpath", value = "//*[@id='ctl00_PlaceHolderMain_m_excelWebRenderer_ewaCtl_sheetContentDiv_Plan1_0_0']/table/tbody")
        source <- webElem$getPageSource()
        tabuahtml <- readHTMLTable(source[[1]],header = T,as.data.frame = T)
        Sys.sleep(1) 
        tabuahtml[[1]]$Mes <- i-1 #Inclui mes no data frame
        tabuahtml[[1]]$Classe <- k #Inclui mes no data frame
        Sys.sleep(1)
        
        #Grava csv
        write.table(tabuahtml[[1]],append = T,sep = ",",file = "tabels.csv")
        Sys.sleep(1)
        
        
        ################ Salva dados1 embaixo
        #Scrol down
        print("Scroll down - Vai!!")
        
        mouse.move(1357,250)
        mouse.click(button = "left",hold = T)
        Sys.sleep(1)
        mouse.move(1357,640)
        mouse.click(button = "left")
        mouse.move(1357,650)
        mouse.click(button = "left")
        Sys.sleep(1)
        mouse.move(1357,640)
        mouse.click(button = "left")
        mouse.move(1357,650)
        mouse.click(button = "left")
        Sys.sleep(1)
        mouse.click(button = "left")
        mouse.click(button = "left")
        mouse.release(button = "left")
        Sys.sleep(1)
        #
        
        
        webElem <- remDr$findElement(using = "xpath", value = "//*[@id='ctl00_PlaceHolderMain_m_excelWebRenderer_ewaCtl_sheetContentDiv_Plan1_0_4']/table/tbody")
        source <- webElem$getPageSource()
        tabuahtml <- readHTMLTable(source[[1]],header = T,as.data.frame = T)
        Sys.sleep(1)
        tabuahtml[[1]]$Mes <- i-1 #Inclui mes no data frame
        tabuahtml[[1]]$Classe <- k #Inclui mes no data frame
        Sys.sleep(1)
        
        #Page up
        
        mouse.move(1357,250)
        mouse.click(button = "left")
        mouse.click(button = "left")
        mouse.click(button = "left")
        mouse.click(button = "left")
        Sys.sleep(1)
        mouse.click(button = "left")
        mouse.click(button = "left")
        mouse.click(button = "left")
        Sys.sleep(1)
        mouse.click(button = "left")
        mouse.click(button = "left")
        
        
        #Grava csv
        print("Salva2")
        write.table(tabuahtml[[1]],append = T,sep = ",",file = "tabels.csv")
        Sys.sleep(1)
        
        print("Reinicia loop")
        
      }
    }
    
  }
  print("Fim")
  
  remDr$close()
  rD$server$stop()
  system("taskkill /im java.exe /f", intern=FALSE, ignore.stdout=FALSE)
  
}

############################Fim
