library(tidyverse)
library(readxl)
library(panelr)
library(reshape)
library(reshape2)
# msci carbon data - wide
# msci historical esg - wide
# refinitiv - wide
# trucost - long
# sp global esg - wide
setwd("C:/Users/mk255125/OneDrive - Teradata/Desktop/ana/ana")
# Step 0: Load all excel files

rm(list=ls())
gc()

file_list <- list.files(pattern = ".xlsx")

df_list <- lapply(file_list, read_excel)

spg <- read.csv("SP GLOBAL - ESG SCORES.csv")

# Step 1: Convert all wides to long

df_list[[4]] <- NULL

# df[[4]] is long
file_list

names(df_list[[2]])[names(df_list[[2]]) == "ISSUER_ISIN"] <- "ISIN"

# names(df_list[[4]])

df_list[[1]][,c(1:5,7,8,9)] <- NULL

df_list[[2]][,c(1:5,7,8)] <- NULL

df_list[[3]][,c(2:6)] <- NULL

df_list[[4]][,c(1,2,5:16)] <- NULL




################### SPG ###############

spg[,c(1:7,9:11)] <- NULL

spg_long <- melt(spg, id.vars = "ISIN")

spg_long$variable <- as.character(spg_long$variable)

spg_long$year <- substr(spg_long$variable, start = nchar(spg_long$variable) - 3,
                        stop = nchar(spg_long$variable))
#
# spg_long <- cast(spg_long, id.vars = c("ISIN", "year"))


spg_esg <- spg[,c(1:9)]
spg_e <- spg[,c(1,10:18)]
spg_s <- spg[,c(1,19:26)]
spg_g <- spg[,c(1,27:35)]

spg_list <- list(spg_esg, spg_e, spg_s, spg_g)

spg_list <- lapply(spg_list, function(x)
  melt(x, id.vars = "ISIN"))

for (i in seq_along(spg_list)){
  spg_list[[i]]$variable <- as.character(spg_list[[i]]$variable)
  spg_list[[i]]$year <-
    substr(spg_list[[i]]$variable, start = nchar(spg_list[[i]]$variable) - 3,
           stop = nchar(spg_list[[i]]$variable))
}


spg_list[[1]]$variable <- NULL
names(spg_list[[1]]) <- c("ISIN", "ESG_spg", "year")

spg_list[[2]]$variable <- NULL
names(spg_list[[2]]) <- c("ISIN", "E_spg", "year")

spg_list[[3]]$variable <- NULL
names(spg_list[[3]]) <- c("ISIN", "S_spg", "year")

spg_list[[4]]$variable <- NULL
names(spg_list[[4]]) <- c("ISIN", "G_spg", "year")


spg_data <- spg_list %>%
  reduce(inner_join, by = c("ISIN", "year"))

## SPG DONE !!

# 2:

names(df_list[[1]])

# df_list[[1]][,c(16:29,58:71,86:99,114,115)] <- NULL

# reshape(df_list[[1]], direction = "long",
#         varying = list(c(2:15), c(16:29), c(30:43), c(44:57),
#                        c(58:71)),
#         v.names = c("scope_1", "scope_12", "scope_2", "scope_3"),
#         timevar = "year")


carbon_emissions <- df_list[[1]]        

names(carbon_emissions)
carbon_emissions_scope1 <- carbon_emissions[,c(1:15)]

carbon_emissions_scope12 <- carbon_emissions[,c(1,30:43)]

carbon_emissions_scope12_inten <- carbon_emissions[,c(1,44:57)]

carbon_emissions_scope2 <- carbon_emissions[,c(1,72:85)]

carbon_emissions_scope3 <- carbon_emissions[,c(1,100:113)]

carbon_emissions <- list(carbon_emissions_scope1, carbon_emissions_scope2,
                         carbon_emissions_scope3, carbon_emissions_scope12,
                         carbon_emissions_scope12_inten)


carbon_emissions <- lapply(carbon_emissions, function(x)
  melt(x, id.vars = "ISIN"))


for (i in seq_along(carbon_emissions)){
  carbon_emissions[[i]]$variable <- as.character(carbon_emissions[[i]]$variable)
  carbon_emissions[[i]]$year <-
    substr(carbon_emissions[[i]]$variable,
           start = nchar(carbon_emissions[[i]]$variable) - 1,
           stop = nchar(carbon_emissions[[i]]$variable))
}


carbon_emissions[[1]]$variable <- NULL
names(carbon_emissions[[1]]) <- c("ISIN", "scope1_msci_carb", "year")

carbon_emissions[[2]]$variable <- NULL
names(carbon_emissions[[2]]) <- c("ISIN", "scope2_msci_carb", "year")

carbon_emissions[[3]]$variable <- NULL
names(carbon_emissions[[3]]) <- c("ISIN", "scope3_msci_carb", "year")

carbon_emissions[[4]]$variable <- NULL
names(carbon_emissions[[4]]) <- c("ISIN", "scope12_msci_carb", "year")

carbon_emissions[[5]]$variable <- NULL
names(carbon_emissions[[5]]) <- c("ISIN", "scope12inten_msci_carb", "year")

fix_year <- function(x){
  x$year <- paste("20", x$year, sep = "")
}

for (i in seq_along(carbon_emissions)){
  carbon_emissions[[i]]$year <- paste("20", carbon_emissions[[i]]$year, sep = "")
}

carbon_data <- carbon_emissions %>% reduce(full_join, by = c("ISIN", "year"))

####### MSCI_letter (df_list[[2]]) #########


esg_letter <- df_list[[2]]

esgscore <- esg_letter[,c(1,14:25)]
envscore <- esg_letter[,c(1,27:38)]


msci <- list(esgscore, envscore)

msci <- lapply(msci, function(x)
  melt(x, id.vars = "ISIN"))

for (i in seq_along(msci)){
  msci[[i]]$variable <- as.character(msci[[i]]$variable)
  msci[[i]]$year <-
    substr(msci[[i]]$variable,
           start = nchar(msci[[i]]$variable) - 3,
           stop = nchar(msci[[i]]$variable))
  msci[[i]]$value <- as.numeric(msci[[i]]$value)
  msci[[i]]$value <- msci[[i]]$value*10
}


msci[[1]]$variable <- NULL
names(msci[[1]]) <- c("ISIN", "esg_msci_hist", "year")

msci[[2]]$variable <- NULL
names(msci[[2]]) <- c("ISIN", "env_msci_hist", "year")

msci_data <- msci %>% reduce(inner_join, by = c("ISIN", "year"))

# msci DONE



###################  TRUCOST ################

trucost <- df_list[[4]]

trucost <- trucost[,c(1:5,9,10,11)]

names(trucost) <- c("ISIN", "year", "carbonscope1_trucost",
                    "carbonscope2_trucost",
                    "carbonscope3_trucost",
                    "carbon_inten_scope1_trucost",
                    "carbon_inten_scope2_trucost",
                    "carbon_inten_scope3_trucost")


names(trucost)[names(trucost) == "Financial Year"] <- "year"

trucost <- trucost %>% filter(year >= 2010 & year <= 2020)

############## DF LIST 3 - refin############
refin <- df_list[[3]]

# sc1, sc2, sc3, sct, sc_1_2_intensity, esg_refin, e_refin

names(refin)

refinsc1 <- refin[,c(1,110:121)]
refinsc2 <- refin[,c(1,122:133)]
refinsc3 <- refin[,c(1,134:145)]
refinsc_t <- refin[,c(1,146:157)]
refinsc_1_2_inten <- refin[,c(1,159:170)]
refin_esg <- refin[,c(1,2:13)]
refin_e <- refin[,c(1,14:25)]

refin_list <- list(refinsc1, refinsc2, refinsc3, refinsc_t, refinsc_1_2_inten,
                   refin_esg, refin_e)

refin_list <- lapply(refin_list, function(x)
  melt(x, id.vars = "ISIN"))

for (i in seq_along(refin_list)){
  refin_list[[i]]$variable <- as.character(refin_list[[i]]$variable)
  refin_list[[i]]$year <-
    substr(refin_list[[i]]$variable,
           start = nchar(refin_list[[i]]$variable) - 3,
           stop = nchar(refin_list[[i]]$variable))
  refin_list[[i]]$value <- as.numeric(refin_list[[i]]$value)
}

for (i in seq_along(refin_list)){
  refin_list[[i]]$variable <- NULL
}

names(refin_list[[1]]) <- c("ISIN", "sc1_refin", "year")
names(refin_list[[2]]) <- c("ISIN", "sc2_refin", "year")
names(refin_list[[3]]) <- c("ISIN", "sc3_refin", "year")
names(refin_list[[4]]) <- c("ISIN", "sct_refin", "year")
names(refin_list[[5]]) <- c("ISIN", "sc12inten_refin", "year")
names(refin_list[[6]]) <- c("ISIN", "refin_esg", "year")
names(refin_list[[7]]) <- c("ISIN", "refin_e", "year")


refin_data <- refin_list %>% reduce(inner_join, by = c("ISIN", "year"))



####################Add control variables#####################
#
# Leverage Ratio
# Market Cap
# ROA
# ROE  
# Current Ratio
# Policy Emissions
# Target Emissions
# Environmental Expenditures
# Investments

names(refin)

refin_roe <- refin[,c(1,339:350)]
refin_roa <- refin[,c(1,351:362)]
refin_lr <- refin[,c(1,375:386)]
refin_mktcap <- refin[,c(1,423:434)]
refin_cr <- refin[,c(1,363:374)]
refin_policyemissions <- refin[,c(1,219:230)]
refin_targetemissions <- refin[,c(1,231:242)]
refin_envexpen <- refin[,c(1,243:254)]

refin_control_list <- list(refin_roe, refin_cr, refin_roa, refin_lr,
                           refin_mktcap, refin_policyemissions,
                           refin_targetemissions, refin_envexpen)

refin_control_list <- lapply(refin_control_list, function(x)
  melt(x, id.vars = "ISIN"))

for (i in seq_along(refin_control_list)){
  refin_control_list[[i]]$variable <- as.character(refin_control_list[[i]]$variable)
  refin_control_list[[i]]$year <- substr(refin_control_list[[i]]$variable,
                                         start = nchar(refin_control_list[[i]]$variable) - 3,
                                         stop = nchar(refin_control_list[[i]]$variable))
}

for (i in seq_along(refin_control_list)){
  refin_control_list[[i]]$variable <- NULL
}


# refin_control_list <- list(refin_roe, refin_cr, refin_roa, refin_lr,
#                            refin_mktcap, refin_policyemissions,
#                            refin_targetemissions, refin_envexpen)

refin_control_list[[1]]

names(refin_control_list[[1]]) <- c("ISIN", "roe", "year")
names(refin_control_list[[2]]) <- c("ISIN", "cr", "year")
names(refin_control_list[[3]]) <- c("ISIN", "roa", "year")
names(refin_control_list[[4]]) <- c("ISIN", "lr", "year")
names(refin_control_list[[5]]) <- c("ISIN", "mktcap", "year")
names(refin_control_list[[6]]) <- c("ISIN", "policyemissions", "year")
names(refin_control_list[[7]]) <- c("ISIN", "targetemissions", "year")
names(refin_control_list[[8]]) <- c("ISIN", "eei", "year")


refin_controls <- refin_control_list %>% reduce(inner_join, by = c("ISIN", "year"))

################# PUT IT ALL TOGETHER ##################

final_list <- list(refin_data, msci_data, spg_data, trucost, carbon_data,
                   refin_controls)

# rm(list = setdiff(ls(), "final_list"))
#
# gc()

for (i in seq_along(final_list)){
  final_list[[i]]$year <- as.character(final_list[[i]]$year)
}

final_dataset <- final_list %>%
  reduce(left_join, by = c("ISIN", "year"))

# Get the company name, country, industry, and continent

names(spg)

spg <- read.csv("SP GLOBAL - ESG SCORES.csv")
spg <- spg[,c(1,4,5,8)]

refin <- read_excel("Refinitiv_Data_Latest.xlsx")

names(refin)

refin_industry <- refin[,c(1,4,5,6)]

final_dataset <- full_join(spg, final_dataset, by = "ISIN")
final_dataset <- full_join(refin_industry, final_dataset, by = "ISIN")

final_dataset <- final_dataset %>%
  group_by(ISIN, year) %>%
  filter(year < 2021) %>%
  arrange(ISIN, year)

# final_dataset <- final_dataset[,c(1,3,2:ncol(final_dataset))]

names(final_dataset)[names(final_dataset) == "Name"] <- "company_name_refin"

# Load company name from msci, carbon, trucost

msci_name <- read_excel("MSCI Carbon Data.xlsx")
msci_name <- msci_name %>% select(ISIN, ISSUER_NAME, ISSUER_CNTRY_DOMICILE)

names(msci_name)

trucost <- read_excel("TRUCOST Carbon Data.xlsx")
names(trucost)
trucost_name <- trucost %>% select(ISIN, Company, Country) %>%
  distinct(ISIN, Company, Country)
names(trucost_name) <- c("ISIN", "company_name_trucost", "country_trucost")


# Add industry
names(final_dataset)[names(final_dataset) == "TRBC Industry Group"] <-
                    "trbc_industry"

final_dataset <- full_join(msci_name, final_dataset, by = "ISIN")

final_dataset <- full_join(trucost_name, final_dataset, by = "ISIN")

final_dataset[,3] <- NULL
## Create a singular name column

final_dataset$full_company_name <- apply(final_dataset[,c("company_name_trucost",
                                                          "company_name_refin",
                                                          "ISSUER_NAME",
                                                          "SP_COMPANY_NAME")], 1, function(x) x[!is.na(x)][1])



final_dataset[,c(2,3,5,7,8,10,23,24)] <- NULL
final_dataset <- final_dataset[,c(1,6,36,2,3,4,5,7:35)]

final_dataset <- final_dataset[,c(1:6,12:17,7:11,18:36)]
#
#
# final_dataset <- final_dataset[,c(1:5,11:16,6:10,17:21,22:34)]

## Rearrange