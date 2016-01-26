#################################################################################################################################
################################################### AssertLife Lapse Analysis ###################################################
#################################################################################################################################

#####################################################################################################
# Setup #
#########

Path <- getwd()
source(paste(Path, "/Initialize.R", sep = ""))

#####################################################################################################
# Load Data #
#############

# Load Correct Claims File Here
Dat <- read.xlsx(paste(Path,"/Data/Lapse_Data.xlsx", sep = ""), 
                 sheet = "Lapse_Data", 
                 detectDates = TRUE, 
                 colNames = TRUE)

Dat <- data.table(Dat)

Increase <- Dat$Increase
Increase[Increase == "NTU"] <- NA
Increase <- as.numeric(Increase)

Dat$Increase <- Increase

colnames(Dat)[which(colnames(Dat) %in% "Smoker/Non-Smoker")]                                        <- "Smoker_Non-Smoker"
colnames(Dat)[which(colnames(Dat) %in% "Adjusted_x000D_.Commencement_x000D_.date")]                 <- "Comm_Dt"
colnames(Dat)[which(colnames(Dat) %in% "Adjusted_x000D_.Voice_x000D_.logged_x000D_.date")]          <- "Voice_Dt"
colnames(Dat)[which(colnames(Dat) %in% "Adjusted_x000D_.status_x000D_.effective_x000D_.end.date")]  <- "End_Dt"
colnames(Dat)[which(colnames(Dat) %in% "Adjusted_x000D_.Status")]                                   <- "Stat"
colnames(Dat)[which(colnames(Dat) %in% "Active.quoted.Premium")]                                    <- "Act_Prem"
colnames(Dat)[which(colnames(Dat) %in% "Total.Amount.paid.since.inception.to.current.month")]       <- "Tot_Pd"
colnames(Dat)[which(colnames(Dat) %in% "Credit.Protection.-.Death.sum.assured")]                    <- "Death_Sum"
colnames(Dat)[which(colnames(Dat) %in% "Credit.Protection.-.Disability.sum.assured")]               <- "Disb_Sum"
colnames(Dat)[which(colnames(Dat) %in% "Credit.Protection.-.Temporary.Disability.monthly.benefit")] <- "TTD_Sum"
colnames(Dat)[which(colnames(Dat) %in% "Critical.Illness.Cover")]                                   <- "CI_Sum"
colnames(Dat)[which(colnames(Dat) %in% "Retrenchment.benefit")]                                     <- "Retr_Sum"
colnames(Dat)[which(colnames(Dat) %in% "RAF.benefit")]                                              <- "RAF_Sum"

str(Dat)
head(Dat)

Dat$Final_Stat <- 0
Dat$Final_Stat[Dat$Stat == "NTU"] <- 1

Dat_Save <- Dat

full.features <- c("Gender", "Age", "Level.of.Income", "Education", "Postal.code", "Product.code", "Agent.name", "Act_Prem",
                   "Death_Sum", "Disb_Sum", "TTD_Sum", "CI_Sum", "Retr_Sum", "RAF_Sum", "Comm_Dt", "Voice_Dt", "Duration",
                   "Final_Stat", "Stat", "Increase")

Dat <- as.data.frame(Dat_Save)[,full.features]
Dat <- Dat[rowSums(is.na(Dat)) < 13, ]

# Dat$Final_Stat <- 0
# Dat$Final_Stat[Dat$Stat == "LAP" & Dat$Duration == 5] <- 1

Dat$Voice_Day <- as.numeric(format(Dat$Voice_Dt, "%d"))
Dat$Voice_Month <- as.numeric(format(Dat$Voice_Dt, "%m"))

Dat$Com_Day <- as.numeric(format(Dat$Comm_Dt, "%d"))
Dat$Com_Month <- as.numeric(format(Dat$Comm_Dt, "%m"))

Dat$rand <- runif(nrow(Dat))

feature.names <- c("Gender", "Age", "Level.of.Income", "Education", "Postal.code", "Product.code", "Agent.name", "Act_Prem",
                   "Death_Sum", "Disb_Sum", "TTD_Sum", "CI_Sum", "Retr_Sum", "RAF_Sum", "Voice_Day", "Voice_Month",
                   "Com_Day", "Com_Month", "Increase")

train <- subset(Dat, rand < 0.8) 
test  <- subset(Dat, rand >= 0.8)

# put IDs for text variables - replace missings by 0 - boosting runs faster
for (f in feature.names) {
  if (class(train[[f]])=="character") {
    levels <- unique(c(train[[f]], test[[f]])) # Combines all possible factors (both from train and test)
    train[[f]] <- ifelse(is.na(train[[f]]), 0, as.integer(factor(train[[f]], levels=levels))) # Create a numeric ID
    test[[f]]  <- ifelse(is.na(test[[f]]), 0, as.integer(factor(test[[f]],  levels=levels)))  # instead of a character ID (If NA then 0)
  }
  else if (class(train[[f]]) == "integer" | class(train[[f]]) == "numeric"){
    train[[f]] <- ifelse(is.na(train[[f]]), 0, train[[f]]) # If NA then
    test[[f]] <- ifelse(is.na(test[[f]]), 0, test[[f]])    # 0
  }
}

tra <- train[,feature.names]

h <- base::sample(nrow(train), 0.2 * nrow(train)) # Takes a sample of 20% of the training data

# Split training data into 2 datasets
dval <- xgb.DMatrix(data = data.matrix(tra[h,]), 
                    label = train$Final_Stat[h])
dtrain <- xgb.DMatrix(data = data.matrix(tra[-h,]), 
                      label = train$Final_Stat[-h])

watchlist <- list(val = dval, train = dtrain)

param <- list(  objective           = "binary:logistic", 
                booster             = "gbtree",
                eta                 = 0.1,
                max_depth           = 8,
                subsample           = 0.7,
                colsample_bytree    = 0.7
)

mod <- xgb.train(   params              = param, 
                    data                = dtrain, 
                    nrounds             = 1000,
                    verbose             = 1, # Whether to print eoutput or not
                    early.stop.round    = 30,
                    watchlist           = watchlist,
                    maximize            = TRUE,
                    eval.metric         = "auc"        
)

cv.res <- xgb.cv(   params              = param, 
                    data                = dtrain, 
                    nrounds             = 100,
                    verbose             = 1, # Whether to print output or not
                    early.stop.round    = 30,
                    watchlist           = watchlist,
                    maximize            = TRUE,
                    eval.metric         = "auc",
                    nfold               = 4
)

test$pred <- predict(mod, data.matrix(test[, feature.names]))

auc(roc(test$pred, as.factor(test$Final_Stat))) # 0.816134

importance_matrix <- xgb.importance(feature.names, model = mod) # check if this goes faster via creation of dump file xgb.dump
xgb.plot.importance(importance_matrix)
xgb.plot.tree(feature_names = feature.names, model = mod, n_first_tree = 1)





















