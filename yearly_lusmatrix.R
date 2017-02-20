#DYNAMIC LUSMATRIX XLS FILE TO YEARLY LUSMATRIX TXT FILES

#Created by Bep Schrammeijer
#July 2016
#Creates lusmatrix files for each year

#packages required
require(openxlsx)

#load excel workbook ls_matrix_country_ls24_origIMAGEd from.../Scenario_SSP1/IME/ME/ME/lus_matrix_ME_ls24.xlsx, call it wb
ls_matrix <- loadWorkbook("Scenario_SSP1/IME/ME/ME/lus_matrix_ME_ls24.xlsx")
#year_0
ls_matrix_0 <- readWorkbook("Scenario_SSP1/IME/ME/ME/lus_matrix_ME_ls24.xlsx", sheet = "dynamic output", 16, colNames = FALSE, rowNames = FALSE, rows = 16:39, cols = 2:5)
ls_matrix_0[,1:3] <- round(ls_matrix_0[,1:3], 0)
ls_matrix_0[,4] <- round(ls_matrix_0[,4], 2)
write.table(ls_matrix_0, "Scenario_SSP1/IME/ME/ME/lusmatrix.0", sep = "\t", row.names = FALSE, col.names = FALSE)
#year_1
ls_matrix_1 <- readWorkbook("Scenario_SSP1/IME/ME/ME/lus_matrix_ME_ls24.xlsx", sheet = "dynamic output", 16, colNames = FALSE, rowNames = FALSE, rows = 16:39, cols = 11:14)
ls_matrix_1[,1:3] <- round(ls_matrix_1[,1:3], 0)
ls_matrix_1[,4] <- round(ls_matrix_1[,4], 2)
write.table(ls_matrix_1, "Scenario_SSP1/IME/ME/ME/lusmatrix.1", sep = "\t", row.names = FALSE, col.names = FALSE)
#year_2
ls_matrix_2 <- readWorkbook("Scenario_SSP1/IME/ME/ME/lus_matrix_ME_ls24.xlsx", sheet = "dynamic output", 16, colNames = FALSE, rowNames = FALSE, rows = 16:39, cols = 20:23)
ls_matrix_2[,1:3] <- round(ls_matrix_2[,1:3], 0)
ls_matrix_2[,4] <- round(ls_matrix_2[,4], 2)
write.table(ls_matrix_2, "Scenario_SSP1/IME/ME/ME/lusmatrix.2", sep = "\t", row.names = FALSE, col.names = FALSE)
#year_3
ls_matrix_3 <- readWorkbook("Scenario_SSP1/IME/ME/ME/lus_matrix_ME_ls24.xlsx", sheet = "dynamic output", 16, colNames = FALSE, rowNames = FALSE, rows = 16:39, cols = 29:32)
ls_matrix_3[,1:3] <- round(ls_matrix_3[,1:3], 0)
ls_matrix_3[,4] <- round(ls_matrix_3[,4], 2)
write.table(ls_matrix_3, "Scenario_SSP1/IME/ME/ME/lusmatrix.3", sep = "\t", row.names = FALSE, col.names = FALSE)
#year_4
ls_matrix_4 <- readWorkbook("Scenario_SSP1/IME/ME/ME/lus_matrix_ME_ls24.xlsx", sheet = "dynamic output", 16, colNames = FALSE, rowNames = FALSE, rows = 16:39, cols = 38:41)
ls_matrix_4[,1:3] <- round(ls_matrix_4[,1:3], 0)
ls_matrix_4[,4] <- round(ls_matrix_4[,4], 2)
write.table(ls_matrix_4, "Scenario_SSP1/IME/ME/ME/lusmatrix.4", sep = "\t", row.names = FALSE, col.names = FALSE)
#year_5
ls_matrix_5 <- readWorkbook("Scenario_SSP1/IME/ME/ME/lus_matrix_ME_ls24.xlsx", sheet = "dynamic output", 16, colNames = FALSE, rowNames = FALSE, rows = 16:39, cols = 47:50)
ls_matrix_5[,1:3] <- round(ls_matrix_5[,1:3], 0)
ls_matrix_5[,4] <- round(ls_matrix_5[,4], 2)
write.table(ls_matrix_5, "Scenario_SSP1/IME/ME/ME/lusmatrix.5", sep = "\t", row.names = FALSE, col.names = FALSE)
#year_6
ls_matrix_6 <- readWorkbook("Scenario_SSP1/IME/ME/ME/lus_matrix_ME_ls24.xlsx", sheet = "dynamic output", 16, colNames = FALSE, rowNames = FALSE, rows = 16:39, cols = 56:59)
ls_matrix_6[,1:3] <- round(ls_matrix_6[,1:3], 0)
ls_matrix_6[,4] <- round(ls_matrix_6[,4], 2)
write.table(ls_matrix_6, "Scenario_SSP1/IME/ME/ME/lusmatrix.6", sep = "\t", row.names = FALSE, col.names = FALSE)
#year_7
ls_matrix_7 <- readWorkbook("Scenario_SSP1/IME/ME/ME/lus_matrix_ME_ls24.xlsx", sheet = "dynamic output", 16, colNames = FALSE, rowNames = FALSE, rows = 16:39, cols = 65:68)
ls_matrix_7[,1:3] <- round(ls_matrix_7[,1:3], 0)
ls_matrix_7[,4] <- round(ls_matrix_7[,4], 2)
write.table(ls_matrix_7, "Scenario_SSP1/IME/ME/ME/lusmatrix.7", sep = "\t", row.names = FALSE, col.names = FALSE)
#year_8
ls_matrix_8 <- readWorkbook("Scenario_SSP1/IME/ME/ME/lus_matrix_ME_ls24.xlsx", sheet = "dynamic output", 16, colNames = FALSE, rowNames = FALSE, rows = 16:39, cols = 74:77)
ls_matrix_8[,1:3] <- round(ls_matrix_8[,1:3], 0)
ls_matrix_8[,4] <- round(ls_matrix_8[,4], 2)
write.table(ls_matrix_8, "Scenario_SSP1/IME/ME/ME/lusmatrix.8", sep = "\t", row.names = FALSE, col.names = FALSE)
#year_9
ls_matrix_9 <- readWorkbook("Scenario_SSP1/IME/ME/ME/lus_matrix_ME_ls24.xlsx", sheet = "dynamic output", 16, colNames = FALSE, rowNames = FALSE, rows = 16:39, cols = 83:86)
ls_matrix_9[,1:3] <- round(ls_matrix_9[,1:3], 0)
ls_matrix_9[,4] <- round(ls_matrix_9[,4], 2)
write.table(ls_matrix_9, "Scenario_SSP1/IME/ME/ME/lusmatrix.9", sep = "\t", row.names = FALSE, col.names = FALSE)
#year_10
ls_matrix_10 <- readWorkbook("Scenario_SSP1/IME/ME/ME/lus_matrix_ME_ls24.xlsx", sheet = "dynamic output", 16, colNames = FALSE, rowNames = FALSE, rows = 16:39, cols = 92:95)
ls_matrix_10[,1:3] <- round(ls_matrix_10[,1:3], 0)
ls_matrix_10[,4] <- round(ls_matrix_10[,4], 2)
write.table(ls_matrix_10, "Scenario_SSP1/IME/ME/ME/lusmatrix.10", sep = "\t", row.names = FALSE, col.names = FALSE)
#year_11
ls_matrix_11 <- readWorkbook("Scenario_SSP1/IME/ME/ME/lus_matrix_ME_ls24.xlsx", sheet = "dynamic output", 16, colNames = FALSE, rowNames = FALSE, rows = 16:39, cols = 101:104)
ls_matrix_11[,1:3] <- round(ls_matrix_11[,1:3], 0)
ls_matrix_11[,4] <- round(ls_matrix_11[,4], 2)
write.table(ls_matrix_11, "Scenario_SSP1/IME/ME/ME/lusmatrix.11", sep = "\t", row.names = FALSE, col.names = FALSE)
#year_12
ls_matrix_12 <- readWorkbook("Scenario_SSP1/IME/ME/ME/lus_matrix_ME_ls24.xlsx", sheet = "dynamic output", 16, colNames = FALSE, rowNames = FALSE, rows = 16:39, cols = 110:113)
ls_matrix_12[,1:3] <- round(ls_matrix_12[,1:3], 0)
ls_matrix_12[,4] <- round(ls_matrix_12[,4], 2)
write.table(ls_matrix_12, "Scenario_SSP1/IME/ME/ME/lusmatrix.12", sep = "\t", row.names = FALSE, col.names = FALSE)
#year_13
ls_matrix_13 <- readWorkbook("Scenario_SSP1/IME/ME/ME/lus_matrix_ME_ls24.xlsx", sheet = "dynamic output", 16, colNames = FALSE, rowNames = FALSE, rows = 16:39, cols = 119:122)
ls_matrix_13[,1:3] <- round(ls_matrix_13[,1:3], 0)
ls_matrix_13[,4] <- round(ls_matrix_13[,4], 2)
write.table(ls_matrix_13, "Scenario_SSP1/IME/ME/ME/lusmatrix.13", sep = "\t", row.names = FALSE, col.names = FALSE)
#year_14
ls_matrix_14 <- readWorkbook("Scenario_SSP1/IME/ME/ME/lus_matrix_ME_ls24.xlsx", sheet = "dynamic output", 16, colNames = FALSE, rowNames = FALSE, rows = 16:39, cols = 128:131)
ls_matrix_14[,1:3] <- round(ls_matrix_14[,1:3], 0)
ls_matrix_14[,4] <- round(ls_matrix_14[,4], 2)
write.table(ls_matrix_14, "Scenario_SSP1/IME/ME/ME/lusmatrix.14", sep = "\t", row.names = FALSE, col.names = FALSE)
#year_15
ls_matrix_15 <- readWorkbook("Scenario_SSP1/IME/ME/ME/lus_matrix_ME_ls24.xlsx", sheet = "dynamic output", 16, colNames = FALSE, rowNames = FALSE, rows = 16:39, cols = 137:140)
ls_matrix_15[,1:3] <- round(ls_matrix_15[,1:3], 0)
ls_matrix_15[,4] <- round(ls_matrix_15[,4], 2)
write.table(ls_matrix_15, "Scenario_SSP1/IME/ME/ME/lusmatrix.15", sep = "\t", row.names = FALSE, col.names = FALSE)
#year_16
ls_matrix_16 <- readWorkbook("Scenario_SSP1/IME/ME/ME/lus_matrix_ME_ls24.xlsx", sheet = "dynamic output", 16, colNames = FALSE, rowNames = FALSE, rows = 16:39, cols = 146:149)
ls_matrix_16[,1:3] <- round(ls_matrix_16[,1:3], 0)
ls_matrix_16[,4] <- round(ls_matrix_16[,4], 2)
write.table(ls_matrix_16, "Scenario_SSP1/IME/ME/ME/lusmatrix.16", sep = "\t", row.names = FALSE, col.names = FALSE)
#year_17
ls_matrix_17 <- readWorkbook("Scenario_SSP1/IME/ME/ME/lus_matrix_ME_ls24.xlsx", sheet = "dynamic output", 16, colNames = FALSE, rowNames = FALSE, rows = 16:39, cols = 155:158)
ls_matrix_17[,1:3] <- round(ls_matrix_17[,1:3], 0)
ls_matrix_17[,4] <- round(ls_matrix_17[,4], 2)
write.table(ls_matrix_17, "Scenario_SSP1/IME/ME/ME/lusmatrix.17", sep = "\t", row.names = FALSE, col.names = FALSE)
#year_18
ls_matrix_18 <- readWorkbook("Scenario_SSP1/IME/ME/ME/lus_matrix_ME_ls24.xlsx", sheet = "dynamic output", 16, colNames = FALSE, rowNames = FALSE, rows = 16:39, cols = 164:167)
ls_matrix_18[,1:3] <- round(ls_matrix_18[,1:3], 0)
ls_matrix_18[,4] <- round(ls_matrix_18[,4], 2)
write.table(ls_matrix_18, "Scenario_SSP1/IME/ME/ME/lusmatrix.18", sep = "\t", row.names = FALSE, col.names = FALSE)
#year_19
ls_matrix_19 <- readWorkbook("Scenario_SSP1/IME/ME/ME/lus_matrix_ME_ls24.xlsx", sheet = "dynamic output", 16, colNames = FALSE, rowNames = FALSE, rows = 16:39, cols = 173:176)
ls_matrix_19[,1:3] <- round(ls_matrix_19[,1:3], 0)
ls_matrix_19[,4] <- round(ls_matrix_19[,4], 2)
write.table(ls_matrix_19, "Scenario_SSP1/IME/ME/ME/lusmatrix.19", sep = "\t", row.names = FALSE, col.names = FALSE)
#year_20
ls_matrix_20 <- readWorkbook("Scenario_SSP1/IME/ME/ME/lus_matrix_ME_ls24.xlsx", sheet = "dynamic output", 16, colNames = FALSE, rowNames = FALSE, rows = 16:39, cols = 182:185)
ls_matrix_20[,1:3] <- round(ls_matrix_20[,1:3], 0)
ls_matrix_20[,4] <- round(ls_matrix_20[,4], 2)
write.table(ls_matrix_20, "Scenario_SSP1/IME/ME/ME/lusmatrix.20", sep = "\t", row.names = FALSE, col.names = FALSE)
#year_21
ls_matrix_21 <- readWorkbook("Scenario_SSP1/IME/ME/ME/lus_matrix_ME_ls24.xlsx", sheet = "dynamic output", 16, colNames = FALSE, rowNames = FALSE, rows = 16:39, cols = 191:194)
ls_matrix_21[,1:3] <- round(ls_matrix_21[,1:3], 0)
ls_matrix_21[,4] <- round(ls_matrix_21[,4], 2)
write.table(ls_matrix_21, "Scenario_SSP1/IME/ME/ME/lusmatrix.21", sep = "\t", row.names = FALSE, col.names = FALSE)
#year_22
ls_matrix_22 <- readWorkbook("Scenario_SSP1/IME/ME/ME/lus_matrix_ME_ls24.xlsx", sheet = "dynamic output", 16, colNames = FALSE, rowNames = FALSE, rows = 16:39, cols = 200:203)
ls_matrix_22[,1:3] <- round(ls_matrix_22[,1:3], 0)
ls_matrix_22[,4] <- round(ls_matrix_22[,4], 2)
write.table(ls_matrix_22, "Scenario_SSP1/IME/ME/ME/lusmatrix.22", sep = "\t", row.names = FALSE, col.names = FALSE)
#year_23
ls_matrix_23 <- readWorkbook("Scenario_SSP1/IME/ME/ME/lus_matrix_ME_ls24.xlsx", sheet = "dynamic output", 16, colNames = FALSE, rowNames = FALSE, rows = 16:39, cols = 209:212)
ls_matrix_23[,1:3] <- round(ls_matrix_23[,1:3], 0)
ls_matrix_23[,4] <- round(ls_matrix_23[,4], 2)
write.table(ls_matrix_23, "Scenario_SSP1/IME/ME/ME/lusmatrix.23", sep = "\t", row.names = FALSE, col.names = FALSE)
#year_24
ls_matrix_24 <- readWorkbook("Scenario_SSP1/IME/ME/ME/lus_matrix_ME_ls24.xlsx", sheet = "dynamic output", 16, colNames = FALSE, rowNames = FALSE, rows = 16:39, cols = 218:221)
ls_matrix_24[,1:3] <- round(ls_matrix_24[,1:3], 0)
ls_matrix_24[,4] <- round(ls_matrix_24[,4], 2)
write.table(ls_matrix_24, "Scenario_SSP1/IME/ME/ME/lusmatrix.24", sep = "\t", row.names = FALSE, col.names = FALSE)
#year_25
ls_matrix_25 <- readWorkbook("Scenario_SSP1/IME/ME/ME/lus_matrix_ME_ls24.xlsx", sheet = "dynamic output", 16, colNames = FALSE, rowNames = FALSE, rows = 16:39, cols = 227:230)
ls_matrix_25[,1:3] <- round(ls_matrix_25[,1:3], 0)
ls_matrix_25[,4] <- round(ls_matrix_25[,4], 2)
write.table(ls_matrix_25, "Scenario_SSP1/IME/ME/ME/lusmatrix.25", sep = "\t", row.names = FALSE, col.names = FALSE)
#year_26
ls_matrix_26 <- readWorkbook("Scenario_SSP1/IME/ME/ME/lus_matrix_ME_ls24.xlsx", sheet = "dynamic output", 16, colNames = FALSE, rowNames = FALSE, rows = 16:39, cols = 236:239)
ls_matrix_26[,1:3] <- round(ls_matrix_26[,1:3], 0)
ls_matrix_26[,4] <- round(ls_matrix_26[,4], 2)
write.table(ls_matrix_26, "Scenario_SSP1/IME/ME/ME/lusmatrix.26", sep = "\t", row.names = FALSE, col.names = FALSE)
#year_27
ls_matrix_27 <- readWorkbook("Scenario_SSP1/IME/ME/ME/lus_matrix_ME_ls24.xlsx", sheet = "dynamic output", 16, colNames = FALSE, rowNames = FALSE, rows = 16:39, cols = 245:248)
ls_matrix_27[,1:3] <- round(ls_matrix_27[,1:3], 0)
ls_matrix_27[,4] <- round(ls_matrix_27[,4], 2)
write.table(ls_matrix_27, "Scenario_SSP1/IME/ME/ME/lusmatrix.27"...
