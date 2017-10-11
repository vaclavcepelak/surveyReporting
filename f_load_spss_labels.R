
# rm(list = ls())
# cat("\014")


#' Installs packages that are not already installed and updates the selected ones
# f_install_if("plotly")

f_install_if <- function(packages = NULL, update = NULL, dependencies = TRUE){
  for (pck in unique(c(packages, update))){
    if (!(pck %in% rownames(installed.packages()))){
      install.packages(pck, dependencies = dependencies, repos = "http://cran.rstudio.com/")
    }
    if (!(pck %in% update)){
      update.packages(pck, dependencies = dependencies, repos = "http://cran.rstudio.com/")
    }
    
    require(pck,character.only = T)
  }
}

f_install_if("data.table", "foreign", "xlsx")


#' Transforms var lab command as a string to list containg information from the command

# com <- commands[81]

f_var_lab <- function(com) {
  scom <- strsplit(com, "\"[^\"\"]*\"(*SKIP)(*F)| ", perl=TRUE)[[1]][-1:-2]
  scom <- matrix(scom, ncol = 2, byrow = TRUE)
  
  return(list(Variable = scom[, 1],
              Label = gsub("[\"\']", "", scom[, 2])))
}

# f_var_lab("var lab age \"Respondent age\"")

f_val_lab <- function(com) {
  scom <- strsplit(com, "\"[^\"\"]*\"(*SKIP)(*F)| ", perl=TRUE)[[1]][-1:-2]
  num <- grep("[\"\']|^[0-9]+$", scom)[1] - 1
  vars <- scom[1:num]
  scom <- scom[-1:-num]
  scom <- matrix(scom, ncol = 2, byrow = TRUE)
  
  labs <- suppressWarnings(if (any(is.na(as.numeric(scom[, 1])))) scom[, 1] else as.numeric(scom[, 1]))
  vals <- gsub("[\"\']", "", scom[, 2])
  return(list(Variable = vars,
              Label = labs,
              Value = vals))
}

# f_val_lab('value labels q1 q2 q3 1 "yes" 2 "no"')


com <- "MRSETS /mcgroup name=$a1 label='label' variables = a1 a2 a3"

f_mrsets(com)

f_mrsets <- function(com) {
  subcoms <- c("name", "label", "variables", "value")
  chunks <- sort(gregexpr(paste(paste0(subcoms, " *="), collapse = "|"), com, ignore.case = TRUE)[[1]])
  chunks <- data.table(start = chunks, stop = c(chunks[-1] - 2, nchar(com)))
  
  chunks <- data.table(chunks = apply(chunks, 1, function(x) substr(com, x[["start"]], x[["stop"]])))
  chunks$type <- tolower(sapply(chunks$chunks, function(x) strsplit(x, " *= *")[[1]][1]))
  chunks$value <- gsub("^ | $", "", sapply(chunks$chunks, function(x) strsplit(x, " *= *")[[1]][2]))
  
  
  return(list(Name = gsub("\\$","",chunks[type == "name", value], ignore.case = TRUE),
              Label = gsub("[\"\']","", chunks[type == "label", value], ignore.case = TRUE),
              Type = factor(gsub(".*m([dc])group.*","\\1", com, ignore.case = TRUE)[1],
                            levels = c("d","c")),
              Variable = strsplit(chunks[type == "variables", value], " ")[[1]],
              Value = if (length(chunks[type == "value", value]) > 0) {
                as.numeric(chunks[type == "value", value])
                }else{NA}))
}


f_spss_if <- function(com) {
  
  
  com <- gsub(" *([=,><]) *","\\1", com) # remove trailing spaces around =,
  
  coms <- strsplit(com, " ")[[1]] # split parts of the command
  cond <- paste(coms[2:(length(coms) - 1)], collapse = " ") # condition part of the command
  cond <- gsub(" and ", " & ", cond) # and change to &
  cond <- gsub(" or ", " | ", cond) # or change to |
  cond <- gsub("not\\(", "!(", cond) # not change to !
  cond <- gsub("=", "==", cond) # use double equal
  cond <- gsub("<>", "!=", cond) # use not equal right
  
  # any
  any_p <- gregexpr("any\\(.*?\\)", cond, ignore.case = TRUE)[[1]]
  
  if(any_p[1] != -1){
  
    any_p <- data.table(start = any_p, stop = any_p + attr(any_p, "match.length") - 1)
    any_all <- data.table(any = apply(any_p, 1, function(x) substr(cond, x[["start"]], x[["stop"]])))
    any_all$type <- suppressWarnings(factor(is.na(as.numeric(gsub("^any\\((.*?),.*$", 
                                                                  "\\1", any_all$any))),
                                            levels = c(FALSE, TRUE), labels = c("vars", "vals")))
    
    # x <- any_all[1,]
    any_all$cond <- apply(any_all, 1, function(x){
      chunks <- strsplit(gsub("any\\((.*)\\)", "\\1", x[["any"]]), ",")[[1]]
      key <- chunks[1]
      chunks <- chunks[2:length(chunks)]
      if (x[["type"]] == "vars"){
        return(paste0("(", paste(paste0(chunks, "==", key), collapse = " | "), ")"))
      }else{
        return(paste0("(", paste(paste0(key, "==", chunks), collapse = " | "), ")"))
      }
    })
    
    for (i in 1:dim(any_all)[1]){
      cond <- gsub(any_all$any[i], any_all$cond[i], cond, fixed = TRUE)
    }
  }
  # range
  range_p <- gregexpr("range\\(.*?\\)", cond, ignore.case = TRUE)[[1]]
  
  if(range_p[1] != -1){
  
    range_p <- data.table(start = range_p, stop = range_p + attr(range_p, "match.length") - 1)
    range_all <- data.table(range = apply(range_p, 1, function(x) substr(cond, x[["start"]], x[["stop"]])))
    range_all$cond <- sapply(range_all$range, function(x){
      chunks <- strsplit(gsub("range\\((.*)\\)", "\\1", x), ",")[[1]]
      return(paste0("(", chunks[1],">=", chunks[2], " & ", chunks[1], "<=", chunks[3], ")"))
    })
    for (i in 1:dim(range_all)[1]){
      cond <- gsub(range_all$range[i], range_all$cond[i], cond, fixed = TRUE)
    }
  }
  
  cond <- gsub("([<>]=)=", "\\1", cond)
  
  # computing part of the command
  
  comp <- coms[length(coms)] 
  comp <- gsub("[\"\']", "", strsplit(comp, "=")[[1]]) # split the variable and value part + remove quotes
  comp[2] <- gsub("$sysmis", "", comp[2]) # $sysmis change to empty string
  
  return(list(Variables = comp[1], Values = comp[2], Conditions = cond))
  
}




#' Takes the vector of spss commands as input
#' Output = factor of command types (specified in the command library)

f_command_library <- function(commands, 
                           command_library = c("var lab", "val lab", "mrsets", "if"),
                           command_functions = list(f_var_lab, f_val_lab, f_mrsets, f_spss_if),
                           short_chars = 3){
  
  command_list <- strsplit(tolower(command_library), " ")
  
  command_type <- factor(sapply(commands, function(command){
    
    c_words <- strsplit(tolower(command), " ")[[1]]
    
    com <- command_list[[1]]
    which_command <- sapply(command_list, function(com){
      k <- length(com)
      if (k == 0) return(NA) else return(all(sapply(1:k, function(w) grepl(com[w], c_words[w]))))
    })
    
    return(command_library[which_command])
  }, USE.NAMES = FALSE),
  levels = command_library)
  
  commands <- commands[!is.na(command_type)]
  command_type <- command_type[!is.na(command_type)]
  
  command_output <- list(command = commands,
                         type = command_type)
  
  command_output$library <- lapply(1:length(commands), function(i) {
    if (command_type[i] %in% command_library) command_functions[[which(command_library == command_type[i])]](commands[i]) else (NA)
    })
  
  return(command_output)
}




# f_command_type(c('VARIABLE LABEL Age "Respondent age"',
#                  'val lab age 999 "NA"',
#                  'MRSETS /mdgroup name=$bank label="Bank client" variables=KBC Barclays Other',
#                  'exe'))






get_commands_metadata <- function(file, outfile = NULL, syntax = paste(readLines(file), collapse = " "),
                                  command_library = c("var lab", "val lab", "mrsets", "if"),
                                  command_functions = list(f_var_lab, f_val_lab, f_mrsets, f_spss_if),
                                  short_chars = 3){
  
  syntax <- gsub(" +", " ", syntax) # replace multiple spaces to one space
  commands <- strsplit(syntax, "\"[^\"\"]*\"(*SKIP)(*F)|\\.", perl = TRUE)[[1]] # split to commands
  commands <- gsub(" +$", "", commands) # trim trailing spaces
  commands <- gsub("^ +", "", commands)
  commands <- commands[sapply(commands, function(x) substr(x, 1, 1)) != "*"] # remove spss comments
  commands <- commands[tolower(commands) != "exe"] # remove exe commands
  commands <- commands[tolower(commands) != ""] # remove empty rows
  
  
  command_output <- f_command_library(commands,
                                      command_library = command_library,
                                      command_functions = command_functions,
                                      short_chars = short_chars)
  if (!is.null(outfile)){
    for(i in 1:length(command_library)){
      command_lib <- command_output$library[command_output$type == command_library[i]]
      if (length(command_lib) > 0){
        command_lib <- do.call("rbind", lapply(command_lib, function(x) as.data.table(x)))
        xlsx::write.xlsx(command_lib, outfile, sheetName = command_library[i], append = (i > 1),
                         row.names = FALSE, showNA = FALSE)
      }  
    }
  }
  
  return(command_output)
}


setwd("C:/Users/Václav/Documents/R skripty/surveyReporting/test_data")
output <- get_commands_metadata("17098-labels-EN.sps")
output <- get_commands_metadata("17098-labels-EN.sps", outfile = "metadata.xlsx")


# data <- read.spss("17098_data.sav", use.value.labels = FALSE, to.data.frame = TRUE)


