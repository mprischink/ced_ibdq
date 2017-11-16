# Installiere Pakete (nur beim ersten Ausfuehren notwendig)
 #install.packages(c("metafor","xlsx"))

# Lade Pakete
library("xlsx")
library("metafor")

###################
## Eingangsdaten ##
###################

# Lese Daten aus Excel Files
dat_pt_all <- read.xlsx("/Users/mprischink/Downloads/sabrina/ced_data.xlsx", sheetName="pt")
dat_pt <- dat_pt_all[c(1:11), ]
dat_pt_mc <- dat_pt_all[c(1,13), ]
dat_pt_cu <- dat_pt_all[c(2:3,12), ]

dat_f_all <- read.xlsx("/Users/mprischink/Downloads/sabrina/ced_data.xlsx", sheetName="f")
dat_f <- dat_f_all[c(1:9), ]
dat_f_mc <- dat_f_all[c(11), ]
dat_f_cu <- dat_f_all[c(1:2,4,10), ]

#################
## Metaanalyse ##
#################

#########################
## Auswertungsfunktion ##
#########################

auswertung <- function(file,yi,vi,data, method="REML"){
  sink(file=file, type="output")
  result <- rma(yi=yi, vi=vi, method=method, data=data, slab=paste(author))
  
  # Ausgabe Gesamteffektstaerke
  print(result)
  
  # Forest-Plot
  forest <- forest(result, cex=1, order=order(result$yi))
  
  # Berechnung Hoehe der ueberschriften ueber Forest-Plot
  y_headers <- nrow(data) +1.5
  
  # ueberschriften Forest-Plot
  text(forest$xlim[1],y_headers,"Autoren und Jahr", pos=4)
  text(forest$xlim[2],y_headers,"ES (95% KI)", pos=2)
  
  # Funnel-Plot
  funnel(result)
  # Labels mit Autoren bei Datenpunkten
  text(result$yi,sqrt(result$vi)-.01,result$slab,cex=0.7)
  
  # Funnel-Plot ohne Labels
  funnel(result)
  
  # Egger-Regressionstest fuer standard metaanalytic model
  result_reg <- regtest (result, model="rma", predictor="vi")
  print(result_reg)
  sink()
}

############
## IG_pt ##
############

auswertung("result_IG_pt.txt",dat_pt$SMD_W_pt_IG,dat_pt$V_W_pt_IG,dat_pt)

############
## IG_f ##
############

auswertung("result_IG_f.txt",dat_f$SMD_W_f_IG,dat_f$V_W_f_IG,dat_f)

############
## KG_pt ##
############

auswertung("result_KG_pt.txt",dat_pt$SMD_W_pt_KG,dat_pt$V_W_pt_KG,dat_pt)

############
## KG_f ##
############

auswertung("result_KG_f.txt",dat_f$SMD_W_f_KG,dat_f$V_W_f_KG,dat_f)

############
## B_pt ##
############

auswertung("result_B_pt.txt",dat_pt$SMD_B_pt,dat_pt$V_B_pt,dat_pt)

############
## B_f ##
############

auswertung("result_B_f.txt",dat_f$SMD_B_f,dat_f$V_B_f,dat_f)

####################################################
## Metaanalysen getrennt M.Crohn/Colitis ulcerosa ##
####################################################

############
## IG_pt_MC ##
############

auswertung("result_IG_pt_MC.txt",dat_pt_mc$SMD_W_pt_IG,dat_pt_mc$V_W_pt_IG,dat_pt_mc,"FE")

############
## IG_pt_CU ##
############

auswertung("result_IG_pt_CU.txt",dat_pt_cu$SMD_W_pt_IG,dat_pt_cu$V_W_pt_IG,dat_pt_cu)

############
## IG_f_MC ##
############

#auswertung("result_IG_f_MC.txt",dat_f_mc$SMD_W_f_IG,dat_f_mc$V_W_f_IG,dat_f_mc,"FE")

############
## IG_f_CU ##
############

auswertung("result_IG_f_CU.txt",dat_f_cu$SMD_W_f_IG,dat_f_cu$V_W_f_IG,dat_f_cu)

############
## KG_pt_MC ##
############

auswertung("result_KG_pt_MC.txt",dat_pt_mc$SMD_W_pt_KG,dat_pt_mc$V_W_pt_KG,dat_pt_mc,"FE")

############
## KG_pt_CU ##
############

auswertung("result_KG_pt_CU.txt",dat_pt_cu$SMD_W_pt_KG,dat_pt_cu$V_W_pt_KG,dat_pt_cu)

############
## KG_f_MC ##
############

#auswertung("result_KG_f_MC.txt",dat_f_mc$SMD_W_f_KG,dat_f_mc$V_W_f_KG,dat_f_mc,"FE")

############
## KG_f_CU ##
############

auswertung("result_KG_f_CU.txt",dat_f_cu$SMD_W_f_KG,dat_f_cu$V_W_f_KG,dat_f_cu)

############
## B_pt_MC ##
############

auswertung("result_B_pt_MC.txt",dat_pt_mc$SMD_B_pt,dat_pt_mc$V_B_pt,dat_pt_mc, "FE")

############
## B_pt_CU ##
############

auswertung("result_B_pt_CU.txt",dat_pt_cu$SMD_B_pt,dat_pt_cu$V_B_pt,dat_pt_cu)

############
## B_f_MC ##
############

#auswertung("result_B_f_MC.txt",dat_f_mc$SMD_B_f,dat_f_mc$V_B_f,dat_f_mc,"FE")

############
## B_f_CU ##
############

auswertung("result_B_f_CU.txt",dat_f_cu$SMD_B_f,dat_f_cu$V_B_f,dat_f_cu)

#############################################
## Resultate mit gewichteten Effektstaerken ##
#############################################

###############
## IG_pt_gew ##
###############

#auswertung("result_IG_pt_gew.txt",dat_pt$gSMD_W_pt_IG,dat_pt$V_W_pt_IG,dat_pt)

##############
## IG_f_gew ##
##############

#auswertung("result_IG_f_gew.txt",dat_f$gSMD_W_f_IG,dat_f$V_W_f_IG,dat_f)

###############
## KG_pt_gew ##
###############

#auswertung("result_KG_pt_gew.txt",dat_pt$gSMD_W_pt_KG,dat_pt$V_W_pt_KG,dat_pt)

##############
## KG_f_gew ##
##############

#auswertung("result_KG_f_gew.txt",dat_f$gSMD_W_f_KG,dat_f$V_W_f_KG,dat_f)

##############
## B_pt_gew ##
##############

#auswertung("result_B_pt_gew.txt",dat_pt$gSMD_B_pt,dat_pt$V_B_pt,dat_pt)

#############
## B_f_gew ##
#############

#auswertung("result_B_f_gew.txt",dat_f$gSMD_B_f,dat_f$V_B_f,dat_f)

#################
## Ersatzcodes ##
#################
#res <- rma(yi=SMD_W_pt_IG, vi=V_W_pt_IG, m1i=M_pre_IG, sd1i=SD_pre_IG, n1i=n_IG_pre, m2i=M_post_IG, sd2i=SD_post_IG, n2i=n_IG_post, method="REML", measure="SMD", data=dat_pt) 
#print(res)

#res1 <- rma(m1i=M_pre_IG, sd1i=SD_pre_IG, n1i=n_IG_pre, m2i=M_post_IG, sd2i=SD_post_IG, n2i=n_IG_post, method="REML", measure="SMD", data=dat_pt) 
#print(res1)


# Ersatz Egger-Regressionstest fuer Model LM
#result_IG_pt_reg <- regtest (result_IG_pt, model="lm")
#print(result_IG_pt_reg)


# order
#result_IG_pt[order(result_IG_pt$yi), ]
#print(result_IG_pt_ordered_by_yi)

closeAllConnections()
