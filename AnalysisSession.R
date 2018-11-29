

##################################################################################################
#                                     ,,                      ,,                
#      db                           `7MM                      db                
#     ;MM:                            MM                                        
#    ,V^MM.    `7MMpMMMb.   ,6"Yb.    MM `7M'   `MF',pP"Ybd `7MM  ,pP"Ybd       
#   ,M  `MM      MM    MM  8)   MM    MM   VA   ,V  8I   `"   MM  8I   `"       
#   AbmmmqMA     MM    MM   ,pm9MM    MM    VA ,V   `YMMMa.   MM  `YMMMa.       
#  A'     VML    MM    MM  8M   MM    MM     VVV    L.   I8   MM  L.   I8       
#.AMA.   .AMMA..JMML  JMML.`Moo9^Yo..JMML.   ,V     M9mmmP' .JMML.M9mmmP'       
#                                           ,V                                  
#                                        OOb"  
#
#                                                                    
#                                             ,,                     
#         .M"""bgd                            db                     
#        ,MI    "Y                                                   
#        `MMb.      .gP"Ya  ,pP"Ybd ,pP"Ybd `7MM  ,pW"Wq.`7MMpMMMb.  
#          `YMMNq. ,M'   Yb 8I   `" 8I   `"   MM 6W'   `Wb MM    MM  
#        .     `MM 8M"""""" `YMMMa. `YMMMa.   MM 8M     M8 MM    MM  
#        Mb     dM YM.    , L.   I8 L.   I8   MM YA.   ,A9 MM    MM  
#        P"Ybmmd"   `Mbmmd' M9mmmP' M9mmmP' .JMML.`Ybmd9'.JMML  JMML.
#                                                                    
####################################################################################################

cat("\nAnalysisSession version: 2018-11-27")
cat("\n\nFirst thing is to load data:")
cat("\n  GetData(\"Proj_####_Data_yyymmdd_hhmm.xlsx\", 0)")
cat("\n\nThen run analyses and make plots and \"Send\" them:")
cat("\n  p1 <- qplot(iris$Sepal.Length, iris$Petal.Length)")
cat("\n  SendPlot(p1)")


## Sys.setenv("R_ZIPCMD" = "C:\\Rtools\\bin\\zip.exe")

## Load various packages ##
library(openxlsx)
library(readxl)
library(httr)
#library(XLConnect)

## ANALYSIS SESSION: Helpers to access data and document analyses -----------------------------


LaunchAnalysisSession <- function( usr, pwd, pk, debug=F)
{
  require(httr)
  if(debug)
  {
      myurl <- paste("http://localhost:60094/DataProject/AnalysisSession.aspx?pk=",pk,"&file="
    , mysession$datafile, sep="")
  }
  else
  {
    myurl <- paste("https://uwac.autism.washington.edu/research/DataProject/AnalysisSession.aspx?pk="
      , mysession$pk,"&file="
      , mysession$datafile, sep="")

  }

  if(pk==0)
  {
    cat("\nLaunching new analysis session...\n")
    cat("\nNEED TO GET THE SESSION PK HERE...\n")
  } else if (pk > 0) {
    cat(paste("\nLoading analysis session #",pk,"...\n"))
    cat(paste("\n\nSession #",mysession$pk, " has been launched.\nAdd results as desired using the SendPlot() & SendOutput() functions.", sep=""))
  }
  
   BROWSE(myurl )
}


GetData <- function(datafile, pk = -1)
{
  require(getPass)
  require(httr)

  usr <- getPass(msg = "Enter your UW NETID: ", noblank = FALSE, forcemask = FALSE)
  pwd <- getPass(msg = "Enter your password: ", noblank = FALSE, forcemask = FALSE)

  # initialize a new environment within which we keep track of indices, etc.
  mysession <- new.env()
  mysession$datafile <- datafile
  mysession$filenamebody <- paste(gsub(".xlsx","", datafile), "_", usr,sep="")
  mysession$user <- usr
  mysession$pwd <- pwd
  mysession$plotnum <- 0
  mysession$tablenum <- 0
  mysession$pk <- pk
  assign("mysession", mysession, envir = .GlobalEnv)

  cat("\nInitiating Analysis Session")
  cat(paste("\nfile: ", datafile))
  cat(paste("\nuser: ", usr,"\n"))


  myurl <- paste("https://uwac.autism.washington.edu//research//stats//DataFileHandler.ashx?file="
  , mysession$datafile, sep="")
  
  r <- GET(myurl 
    , authenticate(user = mysession$user, password = mysession$pwd, type = "ntlm")
    , write_disk(datafile, overwrite=TRUE) )

  datafileinfo <- file.info(datafile)
  
  if(datafileinfo$size > 1000)
  {
    cat("\n\nLoading Excel File...\n\n")
    mysession$xlsx <-  XLinfo(datafile)
    
  
    if(pk >= 0)
    {  
      LaunchAnalysisSession(usr, pwd, pk)
    } 
      
    if(pk==0)
    {
      #Here I need to look up the session pk, number of plots and number of tables
      newpk  <- getPass(msg = "Enter the Session # from the new webpage: ", noblank = FALSE, forcemask = FALSE)
      mysession$pk <- newpk
    }
    
    datarows <- nrow(mysession$xlsx$data[["Data"]])
    datacols <- ncol(mysession$xlsx$data[["Data"]])
    
    nsubj  <- length(unique(mysession$xlsx$data[["Data"]]$id))
    ntimept  <- length(unique(mysession$xlsx$data[["Data"]]$timept))
        
    
    cat(paste("\n\n'Data' worksheet contains: ",datarows, " rows and ", datacols, " variables.", sep=""))
    cat(paste("\n\n'Data' worksheet contains: ",   nsubj, " unique subjects across ", ntimept, " time points.", sep=""))
    cat(paste("\n\nLoad the 'Data' worksheet into a data.frame like this:\n df <- mysession$xldata$data[[\"Data\"]]", sep=""))
    
  }
  else
  {
    return(NA)
  }
  
}


SendPlot <- function(p)
{
  if(mysession$pk == -1)
  {
    cat("\n\nThere is no active Analysis Session.\n\n")
    
  }
  else
    {
    require(mime)
    if(is.ggplot(p))
    {
      mysession$plotnum <- mysession$plotnum + 1
      pname <- paste("PK",mysession$pk ,"_" , mysession$filenamebody,"_PN", mysession$plotnum, sep="")
      
      fil <- paste( pname, ".png", sep="")
      cat(paste("\n\nSaving ggplot object:\n",fil,"\n"))
      
      ggsave(fil, p)
      
      myurl <- "https://uwac.autism.washington.edu/research/Documents/WebDocFileHandler.ashx"
      
      POST(myurl
        , authenticate(user = mysession$user, password = mysession$pwd, type = "ntlm")
        , body = list(y = upload_file(fil)))
      
    }
    else {
      cat("\n\nNot a ggplot object!\n")
    }
  }

}

SendOutput <- function(obj, mytitle = "My Title", stargaze=T)
{
  if(mysession$pk == -1)
  {
    cat("\n\nThere is no active Analysis Session.\n\n")
    
  }
  else
  {
    if(stargaze)
    {
      output <- stargazer(obj, title=mytitle, align=TRUE, type="html", no.space=TRUE)
    } 
    else 
    {
      output <- obj
    }
  
    mysession$tablenum <- mysession$tablenum + 1
    tname <- paste("PK",mysession$pk ,"_" , mysession$filenamebody,"_TN", mysession$tablenum, sep="")
  
    suffix <- ifelse(stargaze==T, ".html", ".txt")
    
    fil <- paste( tname, suffix, sep="")
    cat(paste("\n\nSaving output object:\n",fil,"\n"))
  
    myurl <- "https://uwac.autism.washington.edu/research/Documents/WebDocFileHandler.ashx"

    sink(fil)
    cat(output)
    sink()
  
    POST(myurl
    , authenticate(user = mysession$user, password = mysession$pwd, type = "ntlm")
    , body = list(y = upload_file(fil)))
      
  }
}


## EXCEL: Functions to handle the importing of Excel files and process their contents -----------------------------

 

XLquik <- function(df, path )
{
    dfname <- deparse(substitute(df))
    require(openxlsx)
    q_xl <- createWorkbook("xl")
    addWorksheet(q_xl, dfname)
    writeData(q_xl, sheet=1, df )
    saveWorkbook(q_xl, file=paste(path,"\\",dfname,".xlsx",sep=""), overwrite=T)
}



XLinfo <- function(xlfile, startrow=1, hasheader=TRUE, keepdata=TRUE, mergedata=FALSE
			,onlysheets=NULL, printflag=FALSE, showColUsedInfo=FALSE)
{ 	###################################################################################
	# This function takes a "workbook" object (read from the openxlsx or XLConnect package)
	# and provdies summary information about its contents.
	###################################################################################
	# returns a list object with the following named elements:
	#  sheets (data frame): Info regarding the worksheets in the workbook
	#		- sheet (string)	name of the worksheet
	#		- nrows (numeric)	# of rows of data in the worksheet
	#		- ncols (numeric)	# of columns of data in the worksheet
	###################################################################################
	require(openxlsx)
	#require(XLConnect)
    require(sqldf)
	require(reshape)
    require(readxl)

    last4 <- tolower(substr(xlfile, nchar(xlfile)-3, nchar(xlfile)))
	sheets <- list(); col_names <- list(); data<-list()
	if(printflag) print(paste("##### last4 =", last4))
	
    if (last4 == "xlsx")
    {
		print("Using 'openxlsx' package for .xlsx file")
        if(printflag) print("##### 01a #####")
        wb <- openxlsx::loadWorkbook(xlfile)   	## load the Excel file
		s <- names(wb)							## get list of worksheets

    } else if (last4 == ".xls") {
		print("Using 'readxl' package for .xls file")

		s <- excel_sheets(xlfile)
        #if(printflag) print("##### 01b1 #####")
		#wb <- XLConnect::loadWorkbook(xlfile)   ## load the Excel file
		#s <- getSheets(wb)						## get list of worksheets
	}
    else  {
		if(printflag) print("##### 1 #####")
        print("##ERROR: file not .xlsx or .xls format.")
        return("##ERROR: file not .xlsx or .xls format.")
    }


    if(class(wb)!="Workbook" & class(wb)!="workbook" ) {
        if(printflag) print("##### 2 #####")
        print("##ERROR: input not of class\"workbook\"")
    }
    else {
		if(printflag) cat("#### XLinfo  0 ####\n")
        ### Loop through worksheets
        sheets <- list(); col_names <- list(); data<-list()

		if(printflag) cat("#### XLinfo  0 ####  names(wb):\n")

		if(!is.null(onlysheets))
		{
			s <- s[which(s %in% onlysheets)]
		}

		if(printflag) cat("#### XLinfo  1 ####\n")
        for (i in 1:length(s))
        {
            ## Read the data on the worksheet


			if (last4 == "xlsx") {
				if(printflag) cat("#### XLinfo  1a xlsx ####\n")
				tmp <- readWorkbook(xlfile, sheet=s[i], startRow=startrow )
            } else if (last4 == ".xls") {
				if(printflag) cat("#### XLinfo  1a xls  ####\n")
				#tmp <- XLConnect::readWorksheet(wb, sheet=s[i], startRow=startrow )

				tmp <- read_excel(xlfile, i)

			}

			##tmp <- readWorksheet(wb, sheet = s[i], startRow=startrow, header=hasheader)
			if(printflag) cat("#### XLinfo  1b ####\n")
            if(printflag) cat(paste("s[i]=", s[i],"\n"))
            if(printflag) cat(paste("nrow(tmp)=",nrow(tmp),"\n"))
            if(printflag) cat(paste("ncol(tmp)=", ncol(tmp),"\n"))
            
            if(printflag) print(str(tmp))
            
            tmpnrow <- ifelse(is.null(tmp), 0, nrow(tmp))
            tmpncol <- ifelse(is.null(tmp), 0, ncol(tmp))

            sheets[[i]] <- data.frame(cbind(s[i], tmpnrow, tmpncol), stringsAsFactors=FALSE)
			if(printflag) cat("#### XLinfo  1c ####\n")
            col_names[[i]] <- data.frame(cbind(s[i],names(tmp)), stringsAsFactors=FALSE)
			if(printflag) cat("#### XLinfo  1d ####\n")
            if(tmpnrow>0 | tmpncol>0)
            {
                data[[i]] <- data.frame(tmp, stringsAsFactors=FALSE)
    			if(printflag) cat("#### XLinfo  1e ####\n")
    			if(printflag) cat(paste("#### XLinfo  1e ####  nrows=", nrow(data[[i]]) ,"\n"))
    			if(nrow(data[[i]])>0)
    				{
    			        print(paste("...sheet #", i))
    					data[[i]]$sheet <- s[i]
    					if(printflag) cat("#### XLinfo  1f ####\n")
    					data[[i]]$rownum <- as.numeric(rownames(data[[i]]))  ## create a var to hold the rownumber of the data
    
    					names(data)[i] <- s[i]  ## name the list element
    			}
            }

        }
   		if(printflag) cat("#### XLinfo  2 ####\n")
        if(printflag) cat("#### sheets ####\n")
        if(printflag) print(sheets)
        ## Loop through sheets and drop those with no rows,cols
        
        sheets <- merge_all(sheets)

		names(sheets) <- c("sheet","nrows","ncols")
        sheets$nrows <- as.numeric(sheets$nrows)
        sheets$ncols <- as.numeric(sheets$ncols)
        sheets$sheetnum <- as.numeric(rownames(sheets))

        cols <- merge_all(col_names)
        names(cols) <- c("sheet", "colname")
   		if(printflag) cat("#### XLSXinfo  3 ####\n")


        #######  CHECK COLUMN USAGE ######
        ## Test to see if the column names vary across worksheets
        cols_used <- sqldf("select colname, count(*) as n_used from cols group by colname")

   		if(printflag) cat("#### XLSXinfo  4 ####\n")

		if(length(s) > 1)
		{
			## List columns used fewer than the max:
			maxused <- max(cols_used$n_used)
			n_cols_LTmax <- nrow(cols_used[ which(cols_used$n_used < maxused), ])
			if(n_cols_LTmax > 0)
			{
				if(printflag) cat("#### XLSXinfo  5 ####\n")

			  if(showColUsedInfo)
			  {
    			cat(paste("## CHECK COLUMN USAGE ##","\n"))
  				cat("## NOTICE: The following columns are used less than others.\n")
  				print(cols_used[ which(cols_used$n_used < maxused), ])
			  }
  		}
			else
			{
				if(printflag) cat("#### XLSXinfo  6 ####\n")

				cat("## OK: All columns are used the same number of times.\n")
				if(mergedata)
				{
				mergeddata <- merge_all(data)
				mergeddata <- mergeddata[order(mergeddata$sheet, mergeddata$rownum),]
				}
			}
		}

		if(mergedata)
		{
			if(keepdata)
			{
				if(printflag) cat("#### XLSXinfo  7a ####\n")
				out <- list(sheets = sheets, cols = cols, cols_used = cols_used
					, wbname=xlfile, startrow=startrow, hasheader=hasheader
					, data=data, mergeddata=mergeddata )
			}
			else
			{
				if(printflag) cat("#### XLSXinfo  7b ####\n")
				out <- list(sheets = sheets, cols = cols, cols_used = cols_used
						, wbname=xlfile, startrow=startrow, hasheader=hasheader
						, mergeddata=mergeddata )
			}
		}
		else
		{
			if(keepdata)
			{
				if(printflag) cat("#### XLSXinfo  8a ####\n")
				out <- list(sheets = sheets, cols = cols, cols_used = cols_used
					, wbname=xlfile, startrow=startrow, hasheader=hasheader
					, data=data)
			}
			else
			{
				if(printflag) cat("#### XLSXinfo  8b ####\n")
				out <- list(sheets = sheets, cols = cols, cols_used = cols_used
					, wbname=xlfile, startrow=startrow, hasheader=hasheader )
			}
		}

    }
	class(out) <- "XLinfo"
    return (out)
}



print.XLinfo <- function(xlinfo)
{
        cat("#######################################################\n")
        cat("File Information for: ")
        cat(paste(xlinfo$wbname,"\n"))

        cat(paste(" --> starting row for reading data =", xlinfo$startrow,"\n"))
        cat(paste(" --> assume data includes column header =", xlinfo$hasheader,"\n\n"))

        ####### SUMMARY INFO #######
        cat("### SUMMARY INFO ###\n")
        cat(paste("# of sheets:", nrow(xlinfo$sheets), "\n"))
		if(nrow(xlinfo$sheets) > 1) {
			cat(paste("Min/Max # rows:", min(xlinfo$sheets$nrows), max(xlinfo$sheets$nrows),
				ifelse(min(xlinfo$sheets$nrows) != max(xlinfo$sheets$nrows), "  ==> number of rows vary","") ,"\n"))
			cat(paste("Min/Max # cols:", min(xlinfo$sheets$ncols), max(xlinfo$sheets$ncols),
				ifelse(min(xlinfo$sheets$ncols) != max(xlinfo$sheets$ncols), "  ==> number of columns vary","") ,"\n"))
        } else {
			cat(paste("# rows:", min(xlinfo$sheets$nrows), "\n"))
			cat(paste("# cols:", min(xlinfo$sheets$ncols), "\n"))
		}
		cat("\n")

		####### SHEET INFO #######
        cat("### SHEET INFO ###\n")
		print(xlinfo$sheets[order(as.numeric(rownames(xlinfo$sheets))),])
		cat("\n")

		####### DATA INFO #######
        cat("### DATA INFO ###\n")
		if("data" %in% names(xlinfo))
		{
			for(d in 1: length(xlinfo$data))
			{
			    if(is.na(xlinfo$sheets$ncols[d] )  )
			    {
			        cat("#----------------------------------------------------------------------------------------\n")
			        cat(paste("# SHEET: ", xlinfo$data[[d]]$sheet[1],"\n"))
			        nchars <- nchar(xlinfo$data[[d]]$sheet[1])
			        cat(paste("# -------- NO ROW OR COLUMNS ", collapse=""),"\n",sep="",collapse="")
			    }
			    else if(xlinfo$sheets$ncols[d] == 0  )
			    {
			        cat("#----------------------------------------------------------------------------------------\n")
			        cat(paste("# SHEET: ", xlinfo$data[[d]]$sheet[1],"\n"))
			        nchars <- nchar(xlinfo$data[[d]]$sheet[1])
			        cat(paste("# -------- NO COLUMNS ", collapse=""),"\n",sep="",collapse="")
			    }

                else if(xlinfo$sheets$nrows[d] == 0   )
                {
				cat("#----------------------------------------------------------------------------------------\n")
				cat(paste("# SHEET: ", xlinfo$data[[d]]$sheet[1],"\n"))
				nchars <- nchar(xlinfo$data[[d]]$sheet[1])
				cat(paste("# -------- NO ROWS ", collapse=""),"\n",sep="",collapse="")

				print(head(xlinfo$data[[d]],1))
                }
                else
                {
                cat("#----------------------------------------------------------------------------------------\n")
                cat(paste("# SHEET: ", xlinfo$sheets$sheet[d],"\n"))
                nchars <- nchar(xlinfo$sheets$sheet[d])
                cat(paste("# --------",paste(rep("-",nchars), collapse=""),"\n",sep="",collapse=""))
                print(head(xlinfo$data[[d]],3))

                }
			}
		}
		else
		{
			cat("### ==> NONE: Data from worksheets not retained. ###\n")
		}

		####### MERGED DATA #######
        cat("\n### MERGED DATA INFO ###\n")
		if(nrow(xlinfo$sheets) == 1) {
            cat("### ==> NONE: Only 1 worksheet.")
		} else if(is.null(xlinfo$mergeddata))
        {
            cat("### ==> NONE: Column names vary across worksheets.")
        }
        else
        {
			cat(paste(" # rows:", nrow(xlinfo$mergeddata),"\n"))
            cat(paste(" # cols:", ncol(xlinfo$mergeddata),"  (including new columns: sheet & rownum) \n"))
            cat("# ==> first 3 rows of merged data:\n")
            print(head(xlinfo$mergeddata,3))
            cat("\n\n")
        }
}



loadXLfiles_in_folder <- function(path)
{
### Read the Excel files from a folder into a list
### and save the contents to a text file
require(XLConnect)

xlfiles <- list.files(path)
xlfiles <- xlfiles[grepl(".xls", xlfiles)]

x<-list()

for(i in 1:length(xlfiles)) {
    x[[i]] <- XLinfo(paste(path,xlfiles[[i]], sep=""))
}


sink(paste(path, "XLcontents.txt", sep=""))
#### Display the contents of all the sheets
cat("##################\n")
cat(paste("Excel files from '", path,"'\n",sep=""))
cat("##################\n")

for(i in 1:length(x)) {
    cat(paste("\n\n#################",x[[i]]$wbname,"#################\n"))
    print(x[[i]]$sheets)
    for(j in 1:length(x[[i]]$data)) {
        cat(paste("\n\n# ============",x[[i]]$sheets$sheet[j],"============ #\n",sep=""))
        info(x[[i]]$data[[j]])
    }
}
sink()
cat(paste("\nFile contents saved to: ",path,"XLcontents.txt",sep=""))
return(x)
}


fixXLdatetime <- function(x)
{
  newdate <- as.POSIXct(as.Date(x,origin="1899-12-30"))
  return(newdate)
}

fixXLdate <- function(x)
{
    newdate <- as.Date(x,origin="1899-12-30")
    return (newdate)
}


writeInfoToXL <- function(wb, sh="", df, starting_row=1, starting_col=1)
{
    actual_dfname <- deparse(substitute(df))
    info <- info(df, mode="xl")

    if(sh=="")
        {
        sh <- paste("info",actual_dfname)
        addWorksheet(wb,sh)
    }

    writeTextToXL(wb, sh = sh, starting_row=starting_row, starting_col=starting_col, paste("Summary Info for the Data Frame:", actual_dfname))

    if(info$dfinfo$levelinfo != "")
    {
        writeData(wb, sheet = sh, startRow=starting_row+1, startCol = starting_col, paste("Level Info:",info$dfinfo$levelinfo))
    }
    writeData(wb, sheet=sh, startRow=starting_row+1, startCol = starting_col, paste("N rows:",info$dfinfo$nrows))
    writeData(wb, sheet=sh, startRow=starting_row+1, startCol = starting_col+3, paste("N cols:",info$dfinfo$ncols))

    row_dates <- starting_row + 3
    row_nums <- row_dates + 3 + nrow(info$dates)
    row_txts <- row_dates + 3 + nrow(info$dates) + 3 + nrow(info$nums)

    writeData(wb, sheet = sh, startRow= row_dates, startCol = starting_col, "DATE fields")
    writeDataTable(wb, sheet = sh, startRow= row_dates+1, startCol = starting_col, info$dates, colNames=T, rowNames=F, withFilter=F)

    writeData(wb, sheet = sh, startRow= row_nums, startCol = starting_col, "NUMERIC fields")
    writeDataTable(wb, sheet = sh, startRow= row_nums + 1, startCol = starting_col, info$nums, colNames=T, rowNames=F, withFilter=F)

    writeData(wb, sheet = sh, startRow= row_txts, startCol = starting_col, "TEXT fields")
    writeDataTable(wb, sheet = sh, startRow= row_txts + 1, startCol = starting_col, info$txts, colNames=T, rowNames=F, withFilter=F)

    return(wb)
}


writeInfoByToXL <- function(wb, infoby)
{

    sheet <- paste("info",infoby[[2]]$dfinfo$dfname,"BY",infoby$byvar)
    addWorksheet(wb,sheet)

    for(i in 2:length(infoby))
    {
        if(i==2)
        {
            start_at=1
        }
        else
        {
            start_at = ((i-2) * (infoby[[2]]$dfinfo$ncols + 9 + 8 ))
#            for(j in 2:(i-1))
#            {
#                ##add 9 for the blank rows between each var type
#                starting_row = starting_row + 9 + nrow(infoby[[j]]$dates) + nrow(infoby[[j]]$nums) + nrow(infoby[[j]]$txts)
#            }
        }

        wb = writeInfoToXL(wb, sh=sheet, df=infoby[[i]], starting_row=start_at)
    }

    return(wb)
}



writeInfoByToXL2 <- function(wb, df, byvar)
{
    actual_dfname <- deparse(substitute(df))
    sheet <- paste("info for [",actual_dfname,"] BY",byvar)
    addWorksheet(wb,sheet)

    infobyresults <- infoby(df, byvar, mode="xl")

    for(i in 2:length(infobyresults))
    {
        if(i==2)
        {
            start_at=1
        }
        else
        {
            start_at = ((i-2) * (infobyresults[[i]]$dfinfo$ncols + 9 + 8 ))
#            for(j in 2:(i-1))
#            {
#                ##add 9 for the blank rows between each var type
#                starting_row = starting_row + 9 + nrow(infoby[[j]]$dates) + nrow(infoby[[j]]$nums) + nrow(infoby[[j]]$txts)
#            }
        }

        wb = writeInfoByToXL2_subset(wb, sh=sheet, inforesults=infobyresults[[i]], starting_row=start_at)
    }

    return(wb)
}



writeInfoByToXL2_subset <- function(wb, sh="", inforesults, starting_row=1, starting_col=1)
{
    actual_dfname <- deparse(substitute(df))
    #info <- info(df, mode="xl")

    if(sh=="")
        {
        sh <- paste("info",actual_dfname)
        addWorksheet(wb,sh)
    }

    writeTextToXL(wb, sh = sh, starting_row=starting_row, starting_col=starting_col, paste("Summary Info for the Data Frame:", actual_dfname, inforesults[[1]]$levelinfo))

    if(inforesults$dfinfo$levelinfo != "")
    {
        writeData(wb, sheet = sh, startRow=starting_row+1, startCol = starting_col, paste("Level Info:",inforesults$dfinfo$levelinfo))
    }
    writeData(wb, sheet=sh, startRow=starting_row+1, startCol = starting_col, paste("N rows:",inforesults$dfinfo$nrows))
    writeData(wb, sheet=sh, startRow=starting_row+1, startCol = starting_col+3, paste("N cols:",inforesults$dfinfo$ncols))

    row_dates <- starting_row + 3
    row_nums <- row_dates + 3 + nrow(inforesults$dates)
    row_txts <- row_dates + 3 + nrow(inforesults$dates) + 3 + nrow(inforesults$nums)

    writeData(wb, sheet = sh, startRow= row_dates, startCol = starting_col, "DATE fields")
    writeDataTable(wb, sheet = sh, startRow= row_dates+1, startCol = starting_col, inforesults$dates, colNames=T, rowNames=F, withFilter=F)

    writeData(wb, sheet = sh, startRow= row_nums, startCol = starting_col, "NUMERIC fields")
    writeDataTable(wb, sheet = sh, startRow= row_nums + 1, startCol = starting_col, inforesults$nums, colNames=T, rowNames=F, withFilter=F)

    writeData(wb, sheet = sh, startRow= row_txts, startCol = starting_col, "TEXT fields")
    writeDataTable(wb, sheet = sh, startRow= row_txts + 1, startCol = starting_col, inforesults$txts, colNames=T, rowNames=F, withFilter=F)

    return(wb)
}

writeTextToXL <- function(wb, sh, text, starting_row=1, starting_col=1, hastitle=T)
{

    writeData(wb, sheet = sh, text, startRow=starting_row, startCol = starting_col)

    boldStyle <- createStyle(textDecoration=c("BOLD","UNDERLINE"))

    addStyle(wb, sheet=sh, style=boldStyle,
        rows = starting_row, cols = starting_col, gridExpand = FALSE)

}



XLcols <- function(){
    l2 <- c(" ","A","B","C", "D", "E", "F", "G")
    l1 <- LETTERS #rep(LETTERS, 8)
    l12 <- merge(l1,l2, all=T)
    l12$col <- paste(l12$y,l12$x,sep="")
    l3 <- sqldf("select col from l12 order by col")
    l3$colnum <- seq(1:nrow(l3))
    l3$col <- gsub(" ","",l3$col)
    return(l3)
}

XLcol <- function(x)
{
    if(is.numeric(x))
    {
        L1 <- trunc(x / 26, 0)
        L2 <- x %% 26
        if(L1 >= 1) col <- paste(LETTERS[L1], LETTERS[L2], sep="")
        else col <- LETTERS[L2]
    } else if (is.character(x)) {
        if(nchar(x) == 1)
        {
            col <- which(LETTERS==toupper(x))    
        } else if(nchar(x) == 2) {
            x1 <- substr(x,1,1)
            x2 <- substr(x,2,2)
            
            col <- (which(LETTERS==toupper(x1))*26) + 
                   which(LETTERS==toupper(x2))    
        } else {
            col = "Must be only 1 or 2 characters."
        }
    }
    return(col)
}




XLcolindex <- function(y, debug=F)
{
    foo <- strsplit(y,",")
    if(debug==T) print(paste("foo =",foo))
    colidx <- array()
    counter <- 1
    for(c in 1:length(foo[[1]]))
    {
        if(debug==T) print(paste("foo[c] =",foo[c]))
        hascolon <- grep(":",foo[[1]][c])
        if(debug==T) print(paste("hascolon =", hascolon))
        if(length(hascolon)==0)
        {   # not a range
            colidx[counter] <- XLcolindex_single(foo[[1]][c])
            counter <- counter + 1
        } else {
            newrange <- XLcolindex_range(foo[[1]][c])
            if(debug==T) print("length(colidx)")
            if(debug==T) print(length(colidx))
            if(length(colidx)>1)
            {
                colidx <- c(colidx, newrange) 
            } else {
                colidx <- newrange
            }
            counter <- length(colidx)
        }
        if(debug==T) print(paste(foo[c], colidx[c]))
    }
    return(colidx)
}


XLcolindex_range <- function(x)
{
    foo <- strsplit(x,":")
    idx1 <- XLcolindex_single(foo[[1]][1])
    idx2 <- XLcolindex_single(foo[[1]][2])
    idx_range <- c(idx1 : idx2)
    return (idx_range)
}

XLcolindex_single <- function(x)
{
    idx <- 0
    for(i in 1:nchar(x))
    {
        digits_place <- nchar(x) - i
        letteridx <- grep(tolower(substr(x,i,i)), letters)
        tmpidx <- ifelse( digits_place==0,  letteridx,  (letteridx*(26^digits_place))) #+ letteridx
        idx <- idx + tmpidx
        #print(paste(substr(x,i,i)," i=",i," digits_place=",digits_place," letteridx=",letteridx," tmpidx=",tmpidx," idx=",idx,sep=""))
        
    }
    return (idx)
}

## Utilities ####


# Load Sheets into common dataframe of all
XLsheets_in_xllist <- function(xllist)
{
    path_to_remove <- "C:\\\\Consulting\\\\TMDW\\\\SNC\\\\recd files\\\\"
    datasheets <- NULL
    dupsheets <- NULL
    # Loop through all the XL files in the list
    counter <- 0; dupcounter <- 0;
    for(i in 1:length(xllist))
    {
        tmpsheets <- xllist[[i]]$sheets
        tmpsheets <- tmpsheets[which(tmpsheets$nrows>0),]
        
        #Look through each sheet and get the field names
        # and column sums first and last columns of first and last 10 rows
        # (this integer representation will help check
        #  for duplicate data)
        for(s in 1:nrow(tmpsheets))
        {
            tmp <- tmpsheets[s, ]
            tmp$filename <- gsub(path_to_remove,"",xllist[[i]]$wbname)
            tmp$xllist <- deparse(substitute(xllist))
            tmp$xllistnum <- i
            tmp <- tmp[which(tmp$nrows>0),]

            tmpdata <- xllist[[i]]$data[[tmp$sheet[1]]]
            flds <- names(tmpdata)
            fldscsv <- paste(flds, collapse=",")
            tmp$flds <- fldscsv
            
            tmp$colsums_third <- sum(XLcolumnsum(tmpdata, 3))
            tmp$colsums_last <- sum(XLcolumnsum(tmpdata, ncol(tmpdata)))
            
            #sheets_[[length(sheets_) + 1]] <- tmp
            
            if(counter==0)
            { 
                datasheets <- tmp
                counter <- counter+1
            }
            else
            {
                #Check for dups
                potdups <- sqldf("select *
                    from datasheets a
                    join tmp b ON a.nrows=b.nrows 
                        and a.ncols=b.ncols
                        and a.colsums_third=b.colsums_third
                        and a.colsums_last=b.colsums_last")

                if(nrow(potdups)>=1)
                {
                    # Yes, this is a dup
                    if(dupcounter==0)
                    {
                        dupsheets <- tmp
                        dupcounter <- dupcounter + 1
                    }
                    else
                    {
                        dupsheets <- rbind(dupsheets, tmp)
                        dupcounter <- dupcounter + 1
                    }
                } 
                else 
                {
                    # not a dup
                    datasheets <- rbind(datasheets, tmp)
                    counter <- counter+1
                }
            }
                
        }
    }
    

    out <- list()
    out[[1]] <- datasheets
    if(!is.null(dupsheets)) 
    {
        out[[2]] <- dupsheets
    }
    else {
        out[[2]] <- "No duplicate worksheets"
    }
    names(out) <- c("datasheets", "dupsheets")
    
    return( out )
}


XLcommonsheets <- function(sheetslist, xllist)
{
    # Combine worksheets that contain the same fields
    # Returns a list of compiled dataframes 
    datasheets <- sheetslist$datasheets
    
    datalist <- list()
    commonflds <- unique(datasheets$flds)
    commonflds <- commonflds[which(commonflds!="")]
    
    for(i in 1:length(commonflds))
    {
        cat(paste("\n\n######",i,"\n"))
        print(commonflds[i])
        
        counter <- 1
        for(j in 1:nrow(datasheets))
        {
            #Do the flds match this commonflds?
            if(datasheets$flds[j] == commonflds[i])
            {
                print(paste("  -- match --", counter))
                mysheet <- datasheets$sheet[j]
                myxl <- datasheets$xllistnum[j]
                #print(paste(mysheet, myxl))
                tmpdata <- xllist[[myxl]]$data[[mysheet]]
                tmpdata <- DATEtoTEXT(tmpdata)
                tmpdata$filename <- xllist[[myxl]]$wbname
                
                #print(head(tmpdata,1))
                if(counter==1) {data <- tmpdata}
                else {data <- rbind(data, tmpdata)}
                counter <- counter+1
            }
        }
        datalist[[i]] <- data
    }
    class(datalist) <- "XLdatalist"
    return(datalist)
}


print.XLdatalist <- function(xldatalist){
    
    for(i in 1:length(xldatalist))
    {
        nsheets <- nrow(sqldf())
        cat(paste("\n\n####",i,"####\nnrows=",nrow(xldatalist[[i]])
            ,"  ncols=",ncol(xldatalist[[i]]),"\n\n"))
        print(head(xldatalist[[i]],3))
    }
}


XLcolumnsum <- function(df, columnidx)
{
    if(nrow(df) > 40)
    {
        df <- df[c(1:20, (nrow(df)-20):nrow(df)), ]
    }
    colvals <- sapply(as.character(df[,columnidx]), function(x) sum(strtoi(charToRaw(x),16L)))
    return(colvals)
}



XLsummarize_data_sets <- function(sheets, datalist, path, filename)
{
    wb <- createWorkbook("wb")
    addWorksheet(wb, "XLsheets")
    
    writeData(wb,  "XLsheets", "Data sheets"
        , startRow = 1, startCol = 1)
    writeData(wb,  "XLsheets", sheets$datasheets[ , 1:5]
        , startRow = 2, startCol = 1)
    
    writeData(wb,  "XLsheets", "Duplicate sheets"
        , startRow = nrow(sheets$datasheets) + 5, startCol = 1)
    writeData(wb,  "XLsheets", sheets$dupsheets[ , 1:5]
        , startRow = nrow(sheets$datasheets) + 6, startCol = 1)
    
    
    addWorksheet(wb, "sets samples")
    currentrow <- 3
    for(i in 1:length(datalist))
    {
        tmp <- datalist[[i]]
        tmpfiles <- sqldf("select sheet, filename, count(*) n 
            from tmp group by sheet, filename")
        writeData(wb,  "sets samples", paste("set",i)
            , startRow = currentrow , startCol = 1)
        writeData(wb,  "sets samples", tmpfiles
            , startRow = currentrow , startCol = 2)
        writeData(wb,  "sets samples", head(datalist[[i]], nrow(tmpfiles))
            , startRow = currentrow , startCol = 6)
        currentrow <- currentrow + nrow(tmpfiles) + 2
    }
    for(i in 1:length(datalist))
    {
        infosheet <- paste("set", sprintf("%02.0f",i),sep="")
        addWorksheet(wb, infosheet)
        writeInfoToXL(wb,sh=infosheet, datalist[[i]])
    }
    
    saveWorkbook(wb, file=paste(path,filename,sep=""), overwrite=T)
}



## Formatting #####

styleBOLD <- createStyle(fontColour = "black"
            , halign = "left", valign = "center", textDecoration = "bold", wrapText = F )

styleBOLD2 <- createStyle(fontColour = "black", fgFill="#EEEEEE"
    , halign = "left", valign = "center", textDecoration = "bold", wrapText = T )


styleCURR <- createStyle( numFmt="CURRENCY" )

styleBLUE <- createStyle( fgFill="lightsteelblue1" )
styleGREEN <- createStyle( fgFill="honeydew2" )

styleBLUE1 <- createStyle( fgFill="#c5d9f1" )
styleBLUE2 <- createStyle( fgFill="#dce6f1" )
styleGREEN1 <- createStyle( fgFill="#c4d79b" )
styleORANGE1 <- createStyle( fgFill="#ffcc99" )

styleGREY1 <- createStyle( fgFill="#d9d9d9" )

#c4d79b  olive
#c5d9f1    ltblue1
#dce6f1   ltblue2
#ffcc99   lt orange


help(createStyle)

#colors()[grep("olive",colors())]
#colors()[grep("green",colors())]
#demo("colors")
#help(colors)
#plotCol(nearRcolor("lightblue", "Lab", dist= 30), nrow=16)
#plotCol(nearRcolor("green", "Luv", dist= 90), nrow=12)
#plotCol(nearRcolor("green", "Lab", dist= 90), nrow=12)

## EXAMPLES ##
#addStyle(wb, 1, styleCURR, rows=c(1:66), cols = currency_cols, gridExpand = TRUE, stack=TRUE )
#setColWidths(wb, 1, c(6:7,25:54), widths = 12)
#setColWidths(wb, sheet = 2, cols = 1:5, widths = "auto")

# XLcol("AA"):XLcol("BB")
# XLcolindex("AA:AF")



