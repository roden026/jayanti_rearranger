# Rearrange csv files for Jayanti

# Install packages. You only need to run this once per computer. It won't hurt anything if
# you run it again, it just takes a little time.
install.packages("openxlsx")
install.packages("stringr")

# Load packages. You only need to run this once per R session. It won't hurt anything if
# you run it again, it just takes a little time.
library(openxlsx)
library(stringr)

#### INPUTS - set these values. 
# Filepath to folder that holds 1 folder for each Rep
# Note the slashes are going the opposite direction that they normally do on a windows machine
headfolder <- '/Users/ericroden/Downloads/Jayanti/'

# xlsx file output name. Needs the .xlsx file extension
outputName <- 'testOutput.xlsx'

# The headfolder will be used to read in the CSVs
# NOTE: THIS GETS ALL OF THE CSV FILES IN THE SUB-FOLDERS 
# SO MAKE SURE THE ONLY CSV FILES IN THE SUB-FOLDERS ARE THE ONES YOU
# WANT TO READ.

# set and store max rt value and min rt value
rt.min <- 11.625
rt.max <- 11.715

####

#### Load these functions

# works with arranged to write the output files
write.arranged <- function(wb, dirsList, rt.min, rt.max){
  
  for(i in 1:length(dirsList)){
    
    # loops over each directory in the head folder
    addWorksheet(wb, str_sub(dirsList[i], 3))
    
    setwd(dirsList[i])
    
    fileList <- list.files()[grepl(".CSV", list.files())]
    
    # loop through each file
    for(j in 1:length(fileList)){
      
      # arrange each file
      arranged <- arranger(fileList[j], rt.min, rt.max)
      
      # set up top row
      sampleNumber <- paste0('Sample', j, '-', i)
      padding <- rep(NA, ncol(arranged)-1)
      topRow <- c(sampleNumber, padding)
      
      # set up data for binding
      topRow2 <- data.frame(t(topRow))
      colnames(topRow2) <- colnames(arranged)
      
      # Attach header row
      arranged <- rbind(topRow2, arranged)
      
      # write data to workbook. 1 sheet per rep, with a spacer column
      writeData(wb, i, arranged, startCol = (((j-1)*(ncol(arranged)+1))+1), colNames = FALSE)

    }
    
    # reset wd
    setwd(headfolder)
    
  }
}

# Takes a filename, rt.min, and rt.max and rearranges it. The script then writes a csv using the same 
# filename with 'Arranger-' added to the front. 
arranger <- function(filename, rt.min, rt.max){
  
  # read in the csv
  data1 <- read.csv(filename, stringsAsFactors = FALSE)
  
  # knock off the unneeded data at the top and remove the excess columns
  data1 <- data1[2:nrow(data1), c(1,2)]
  
  # Find the locations where the columns should split
  cutpoints <- which(grepl("Ion", data1[,1]))
  
  # Add the end point
  cutpoints <- append(cutpoints, nrow(data1)+1)

  # list of ions
  ionNames <- data1[grepl('Ion', data1[ ,1]), 1]
  
  # Set up column names
  ionNames <- str_sub(ionNames, end = ((gregexpr(pattern = ' ', ionNames)[[1]][2])-1))
  colHeads <- c('RT', ionNames)
  
  # make a dataframe and store the first two columns cut at the appropriate location
  df <-  data.frame(data1[cutpoints[1]:(cutpoints[2]-1), 1:2])
  
  # running column names list to avoid issues with indexing/column naming
  columnNames <- c('rt', 'peak1')
  
  # for each additional cut point, cut the data and place it in new columns.
  for(i in 2:(length(cutpoints)-1)) {
    
    # generate new column names
    # I think this should be okay... 
    # columnNames <- append(columnNames, paste0('rt',i))
    columnNames <- append(columnNames, paste0('peak',i))
    
    ### Add in code to add a blank column (df[,'colname' or index] <- NA)
    
    # add new columns
    # This fixes the need to only have the rt once
    # df <- transform(df, newColumn=data1[cutpoints[i]:(cutpoints[i+1]-1), 1])
    df <- transform(df, newColumn=data1[cutpoints[i]:(cutpoints[i+1]-1), 2])
    
    # rename columns
    colnames(df) <-columnNames 
    
  }
  
  
  # find the right data (right rt window)
  returned <- df[df$rt >= rt.min & df$rt <= rt.max, ]
  
  # set up data for binding
  colHeads2 <- data.frame(t(colHeads))
  colnames(colHeads2) <- colnames(returned)
  
  # Attach header row
  returnFrame <- rbind(colHeads2, returned)
  
  #return the data. 
  returnFrame
}
####


#### Run this section
# set directory to containing folder
setwd(headfolder)

# get list of folders contained within the head folder
dirsList <- list.dirs(recursive = FALSE)

# NOTE: THIS SCRIPT WORKS BY GETTING ALL OF THE CSV FILES IN THE SUBFOLDERS (BELOW THE HEADFOLDER)
# SO MAKE SURE THE ONLY CSV FILES IN THOSE FOLDERS ARE THE ONES YOU WANT TO READ.

# create an xlsx workbook
wb <- createWorkbook()

# write your files
write.arranged(wb, dirsList = dirsList, rt.min, rt.max)
# No file has been saved yet, you must run the last two commands to do so.

# save the newly completed workbook. note: this will overwrite any existing workbook with the same name.
setwd(headfolder)
saveWorkbook(wb, file = outputName, overwrite = TRUE)

