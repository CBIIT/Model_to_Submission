#!/usr/bin/env Rscript

#Model_to_Submission.R


##################
#
# USAGE
#
##################

#This script takes the three files that make a CBIIT data model: model, properties, and terms, and creates a submission workbook with formatting and enumerated drop down menus.

#Run the following command in a terminal where R is installed for help.

#Rscript --vanilla Model_to_Submission.R --help

##################
#
# Env. Setup
#
##################

#List of needed packages
list_of_packages=c("dplyr","yaml","stringi","openxlsx","optparse","tools")

#Based on the packages that are present, install ones that are required.
new.packages <- list_of_packages[!(list_of_packages %in% installed.packages()[,"Package"])]
suppressMessages(if(length(new.packages)) install.packages(new.packages, repos = "http://cran.us.r-project.org"))

#Load libraries.
suppressMessages(library(dplyr,verbose = F))
suppressMessages(library(yaml,verbose = F))
suppressMessages(library(stringi,verbose = F))
suppressMessages(library(openxlsx,verbose = F))
suppressMessages(library(optparse,verbose = F))
suppressMessages(library(tools,verbose = F))

#remove objects that are no longer used.
rm(list_of_packages)
rm(new.packages)


##################
#
# Arg parse
#
##################

#Option list for arg parse
option_list = list(
  make_option(c("-m", "--model"), type="character", default=NULL, 
              help="Model file yaml", metavar="character"),
  make_option(c("-p", "--property"), type="character", default=NULL, 
              help="Model property file yaml", metavar="character"),
  make_option(c("-t", "--terms"), type="character", default=NULL, 
              help="Model terms file yaml", metavar="character"),
  make_option(c("-r", "--readme"), type="character", default=NULL, 
              help="README xlsx page (optional)", metavar="character")
)


#create list of options and values for file input
opt_parser = OptionParser(option_list=option_list, description = "\nModel_to_Submission.R v2.0.3\n\nThis script takes the three files that make a CBIIT data model: model, properties, and terms, and creates a submission workbook with formatting and enumerated drop down menus.\n")
opt = parse_args(opt_parser)

#If no options are presented, return --help, stop and print the following message.
if (is.null(opt$model)&is.null(opt$property)&is.null(opt$terms)){
  print_help(opt_parser)
  cat("Please supply all files for the data model, model, property and terms.\n\n")
  suppressMessages(stop(call.=FALSE))
}

#Data model pathway
model_path=file_path_as_absolute(opt$model)

#Data property pathway
property_path=file_path_as_absolute(opt$property)

#Data terms pathway
terms_path=file_path_as_absolute(opt$terms)

#Data readme pathway
if (!is.null(opt$readme)){
  readme_path=file_path_as_absolute(opt$readme)
}

#A start message for the user that the validation is underway.
cat("The submission template file is being created at this time.\n")

###############
#
# Read in files
#
###############

if (!is.null(opt$readme)){
  readme=read.xlsx(xlsxFile =  readme_path,sheet = 1,colNames = F)
}


model=read_yaml(file = model_path)
model_props=read_yaml(file = property_path)
model_terms=read_yaml(file = terms_path)

###########
#
# File name rework
#
###########

#Rework the file path to obtain a file name, this will be used for the output file.
file_name=stri_reverse(stri_split_fixed(stri_reverse(basename(model_path)),pattern = ".", n=2)[[1]][2])

path=paste(dirname(model_path),"/",sep = "")

#Output file.
output_file=paste(file_name,
                  "_CCDITemplate",
                  stri_replace_all_fixed(
                    str = Sys.Date(),
                    pattern = "-",
                    replacement = ""),
                  sep="")


#################
#
# Rework of data frames
#
#################

#Create a preferred order of nodes. While hard coded, any additions to the data model will result with those nodes being placed at the end of the list.
preferred_order=c("study","study_admin","study_arm","study_funding","study_personnel","publication","participant","diagnosis","therapeutic_procedure","sample","sequencing_file","methylation_array_file","imaging_file","clinical_measure_file","synonym")
preferred_order=preferred_order[preferred_order %in% names(model$Nodes)]
preferred_order=c(preferred_order,names(model$Nodes)[!names(model$Nodes) %in% preferred_order])

#Create new Dictionary page
dd=data.frame(matrix(ncol = 9,nrow=0))
dd_add=data.frame(matrix(ncol = 9,nrow=1))

colnames(dd)<-c("Property","Description","Node","Type","Example value","Required","Key","CDE","CDE version")
colnames(dd_add)<-c("Property","Description","Node","Type","Example value","Required","Key","CDE","CDE version")

#Populate Dictionary page
for (prop in names(model_props$PropDefinitions)){
  type_test=model_props["PropDefinitions"][[1]][prop][[1]]["Type"][[1]]
  enum_test=model_props["PropDefinitions"][[1]][prop][[1]]["Enum"][[1]]
  
  dd_add$Property=prop
  dd_add$Description=model_props["PropDefinitions"][[1]][prop][[1]]["Desc"][[1]]
  
  if (!is.null(type_test)){
    dd_add$Type=paste(unlist(model_props["PropDefinitions"][[1]][prop][[1]]["Type"],recursive = T,use.names = F),collapse = ";")
  }else{
    dd_add$Type=NA
  }
  #Checks for enumerated values and then creates a partial list for the data dictionary page.
  if (!is.null(enum_test)){
    dd_add$Type="enum"
    enums=unlist(model_props["PropDefinitions"][[1]][prop][[1]]["Enum"],recursive = T,use.names = F)
    if (length(enums)>4){
      dd_add$`Example value`=paste(paste(enums[1:4],collapse = ";"),";etc (see Terms and Values Sets)",sep="")
    }else{
      dd_add$`Example value`=paste(enums,collapse = ";")
    }
  }else{
    dd_add$`Example value`=NA
  }
  
  #Add the required marker
  if (is.null(model_props["PropDefinitions"][[1]][prop][[1]]["Req"][[1]])){
    dd_add$Required=NA
  }else{
    dd_add$Required=model_props["PropDefinitions"][[1]][prop][[1]]["Req"][[1]]
  }
  
  #Add the Key marker
  if (is.null(model_props["PropDefinitions"][[1]][prop][[1]]["Key"][[1]])){
    dd_add$Key=NA
  }else{
    dd_add$Key=model_props["PropDefinitions"][[1]][prop][[1]]["Key"][[1]]
  }
  
  dd=rbind(dd,dd_add)
}


#Insert source ids for the properties.
df_prop_code=data.frame(matrix(ncol = 4,nrow=0))
df_prop_code_add=data.frame(matrix(ncol = 4,nrow=1))

colnames(df_prop_code)<-c("Property","Code","Version","Source")
colnames(df_prop_code_add)<-c("Property","Code","Version","Source")


#Create list of properties and their CDE codes
for (x in 1:length(names(model_props$PropDefinitions))){
  
  if(any(names(model_props$PropDefinitions[[x]])%in%"Term")){
    if (any(grepl(pattern = "caDSR", x = model_props$PropDefinitions[[x]]["Term"][[1]]))){
      num_codes=grep(pattern = "caDSR", x = model_props$PropDefinitions[[x]]["Term"][[1]])
      for (y in num_codes){
        df_prop_code_add$Property=names(model_props$PropDefinitions[x])
        df_prop_code_add$Code=model_props$PropDefinitions[[x]]["Term"][[1]][[y]]["Code"][[1]]
        df_prop_code_add$Version=model_terms$Terms[names(model_props$PropDefinitions[x])][[1]]["Version"][[1]]
        df_prop_code_add$Source="caDSR"
        df_prop_code=rbind(df_prop_code, df_prop_code_add)
      }
    }
  }
}


#For only caDSR and NCIt sources at this time, creates the columns that notes what the ids are for each property with these values.
for (prop in 1:dim(dd)[1]){
  code=NA
  if (dd$Property[prop] %in% df_prop_code$Property){
    prop_df=filter(df_prop_code, Property==dd$Property[prop])
    if (any(grepl(pattern = "caDSR", x = prop_df$Source))){
      prop_df_caDSR=filter(prop_df,Source=="caDSR")
      if (!is.null(prop_df_caDSR$Code)){
        for (code_pos in 1: dim(prop_df_caDSR)[1]){
          code=prop_df_caDSR$Code[code_pos]
          ver=prop_df_caDSR$Version[code_pos]
          if (is.na(dd$CDE[prop])){
            dd$CDE[prop]= code
            dd$`CDE version`[prop]= ver
          }
        }
      }
    }
  }
}

#Fill out the node column in the DD
for (prop in 1:length(dd$Property)){
  for (node in names(model["Nodes"][[1]])){
    if (dd$Property[prop] %in% model["Nodes"][[1]][node][[1]][[1]]){
      dd$Node[prop]=node
    }
  }
}

#For required properties, place the node value in the required column. (This might change, now that we have a sectioned off node per page, instead of one flat data frame.) 
dd$Required[grepl(pattern = "FALSE",x = dd$Required)]<-NA

dd$Required[grep(pattern = "TRUE",x = dd$Required)]<-dd$Node[grep(pattern = "TRUE",x = dd$Required)]

dd=dd[order(match(dd$Node,preferred_order)),]

#Set up the Terms and Value Set sheet
TaVS=data.frame(matrix(ncol = 4,nrow=0))
TaVS_add=data.frame(matrix(ncol = 4,nrow=1))

colnames(TaVS)<-c("Value Set Name","(subset)","Term","Definition")
colnames(TaVS_add)<-c("Value Set Name","(subset)","Term","Definition")

#Take properties and apply them to a data frame that mirrors TaVS
for (node in preferred_order){
  for (prop in model["Nodes"][[1]][[node]][1][["Props"]]){
    enum_list=model_props["PropDefinitions"][[1]][[prop]]["Enum"][[1]]
    if (!is.null(enum_list)){
      if (length(enum_list)>1){
        if(!prop%in%TaVS$`Value Set Name`){
          for (enum in enum_list){
            
            TaVS_add$`Value Set Name`=prop
            TaVS_add$Term=enum
            if (enum %in% names(model_terms$Terms)){
              TaVS_add$Definition=model_terms$Terms[enum][[1]]['Definition'][[1]]
            }
            TaVS=rbind(TaVS,TaVS_add)
            TaVS_add$Definition=NA
          }
          TaVS_add$Term=NA
          TaVS_add$Definition=NA
          TaVS_add$`Value Set Name`=NA
          TaVS=rbind(TaVS,TaVS_add)
        }
      }
    }
  }
}


#####################
#
# Write out with formatting
#
#####################

#Create workbook
wb = createWorkbook()

if (!is.null(opt$readme)){
  addWorksheet(wb,"README and INSTRUCTIONS")
}

for (node in preferred_order){
  addWorksheet(wb,node)
}

addWorksheet(wb,"Dictionary")
addWorksheet(wb,"Terms and Value Sets")

#Metadata page styles
node_style=createStyle(fontColour = "black", fgFill = "#E7E6E6", textDecoration = "Italic")
prop_style=createStyle(fontColour = "#595959", fgFill = "white")
prop_require_style=createStyle(fontColour = "black",fgFill = "#FFF2CC" , textDecoration = "Bold")
text_format=createStyle(numFmt = "TEXT")

#Dictionary page styles
dd_header_style=createStyle(fontColour = "white",fgFill = "black")

#TaVS page styles
tavs_1=createStyle(fontColour = "black",fgFill = "#DEE6F0")
tavs_2=createStyle(fontColour = "black",fgFill = "#EEDDDC")
tavs_blank=createStyle(fontColour= "black", fgFill = "#FFFFFF")

#Write readme page
if (!is.null(opt$readme)){
  writeData(wb = wb,sheet = "README and INSTRUCTIONS",x = readme,colNames = FALSE)
}

#Insert the key property linking for each node.
for (node in preferred_order){
  metadata=data.frame()
  running_vec=c()
  props=model["Nodes"][[1]][node][[1]]["Props"][[1]]
  
  all_relationships=unlist(model$Relationships)
  node_relationships=names(all_relationships[grep(pattern = node, x = all_relationships)])
  relationships=node_relationships[grep(pattern = ".Src", x = node_relationships)]
  relationships=unique(stri_split_fixed(str = relationships,pattern = ".",n=2,simplify = T)[,1])
  
  for (relation in relationships){
    #Determine key that needs to be used to connect this node to a parent node
    Ends=unlist(model$Relationships[relation][[1]]["Ends"])
    Src=unique(Ends[grep(pattern = ".Src", x = names(Ends))])
    Dsts=Ends[grep(pattern = ".Dst", x = names(Ends))]
    if (Src == node){
      for (Dst in Dsts){
        nodeprops=model["Nodes"][[1]][Dst][[1]]["Props"][[1]]
        for (nodeprop in nodeprops){
          if (!is.null(model_props$PropDefinitions[nodeprop][[1]]$Key)){
            props=c(paste(Dst,nodeprop,sep = "."),props)
          }
        }
      }
    }
  }
  
  #Add the type property for submission and then make the data frame
  props=c("type",props)
  props=unique(props)
  metadata_length=length(props)
  metadata=data.frame(matrix(ncol=metadata_length, nrow=1))
  colnames(metadata)<-props
  metadata$type=node
  
  #Write out each node
  writeData(wb = wb, sheet = node, x=metadata)
  
  #Metadata sheets style apply
  for (col in 1:dim(metadata)[2]){
    if (colnames(metadata[col])=="type" | grepl(pattern = "\\.", x = colnames(metadata[col]))){
      writeData(wb = wb,sheet = node,x = metadata[col], headerStyle = prop_require_style, startCol = col)
    }else if (colnames(metadata[col])%in%dd$Property[!is.na(dd$Required)]){
      writeData(wb = wb,sheet = node,x = metadata[col], headerStyle = prop_require_style, startCol = col)
    }else{
      writeData(wb = wb,sheet = node,x = metadata[col], headerStyle = prop_style, startCol = col)
    }
  }
  
  #Data validation (drop down lists)
  
  #Pull out the positions where the value set names are located
  VSN_new=c(1)
  VSN_counter=1
  while(VSN_counter!=dim(TaVS)[1]){
    if (is.na(TaVS$`Value Set Name`[VSN_counter]) & !is.na(TaVS$`Value Set Name`[(VSN_counter+1)])){
      VSN_new=c(VSN_new,(VSN_counter+1))
    }
    VSN_counter=VSN_counter+1
  }
  
  VSN=VSN_new
  #VSN=grep(pattern = FALSE, x = is.na(TaVS$`Value Set Name`))  
  
  
  #for each instance of a value_set_name, note the position on the Terms and Value Sets page, create a list for each with all accepted values.
  for (prop in props){
    if (prop %in% unique(TaVS$`Value Set Name`)){
      start_pos=min(grep(pattern = prop,x = TaVS$`Value Set Name`))+1
      stop_pos=VSN[start_pos<VSN][1]-1
      #Create restricted drop downs based on the TaVS tab.
      if (start_pos==1){
        col_pos=grep(pattern = TRUE, x = (colnames(metadata) %in% prop))
        suppressWarnings(dataValidation(wb = wb, sheet = node, cols= col_pos,rows = 2:50000, type="list",value = paste("'Terms and Value Sets'!$C$",start_pos,":$C$",stop_pos+1,sep="")))
      }else if (!is.na(stop_pos)){
        col_pos=grep(pattern = TRUE, x = (colnames(metadata) %in% prop))
        suppressWarnings(dataValidation(wb = wb, sheet = node, cols= col_pos,rows = 2:50000, type="list",value = paste("'Terms and Value Sets'!$C$",start_pos,":$C$",stop_pos,sep="")))
      }else{
        col_pos=grep(pattern = TRUE, x = (colnames(metadata) %in% prop))
        suppressWarnings(dataValidation(wb = wb, sheet = node, cols= col_pos,rows = 2:50000, type="list",value = paste("'Terms and Value Sets'!$C$",start_pos,":$C$",dim(TaVS)[1],sep="")))
      }
    }
  }
  addStyle(wb = wb, sheet = node, style = text_format,rows = 2:50000, cols = 1:length(props), gridExpand = TRUE)
}

#Setup boundaries and shifting window for groupRows feature
# VSN=grep(pattern = FALSE, x = is.na(TaVS$`Value Set Name`))
TaVS_row_count=0

for (VSNn in 1:length(VSN)){
  start=VSN[VSNn]+1
  stop=VSN[VSNn+1]-1
  
  #Add color to rows so they alternate on the TaVS page.
  TaVS_row_count=TaVS_row_count+1
  if ((TaVS_row_count %% 2) == 0){
    tavs_style = tavs_1
  }else{
    tavs_style = tavs_2
  }
  #coloring of rows
  if (!is.na(stop_pos)){
    addStyle(wb = wb,sheet = "Terms and Value Sets",style = tavs_style,cols = c(1:4),rows = c((start):(dim(TaVS)[1])), gridExpand = TRUE)
  }else{
    addStyle(wb = wb,sheet = "Terms and Value Sets",style = tavs_style,cols = c(1:4),rows = c(start:dim(TaVS)[1]), gridExpand = TRUE)
  }
  
  #grouping of rows
  if (!is.na(stop)){
    groupRows(wb = wb,sheet = "Terms and Value Sets",rows = c((start+1):(stop)),hidden = TRUE)
  }else{
    groupRows(wb = wb,sheet = "Terms and Value Sets",rows = c((start+1):(dim(TaVS)[1])),hidden = TRUE)
  }
}

#Write out the dictionary and TaVS pages
writeData(wb = wb,sheet = "Dictionary",x = dd)
writeData(wb = wb,sheet = "Terms and Value Sets",x = TaVS)

#Format Dictionary
addStyle(wb = wb,sheet = "Dictionary",style = prop_style,cols = 1,rows = 1:dim(dd)[1]+1)

for (row in 1:dim(dd)[1]){
  if (dd[row,1]%in%dd$Property[!is.na(dd$Required)]){
    addStyle(wb = wb,sheet = "Dictionary",style = prop_require_style,cols = 1,rows = row+1)
  }
}

writeData(wb = wb,sheet = "Dictionary",x = dd[1,], startRow = 1, headerStyle = dd_header_style)

writeData(wb = wb,sheet = "Terms and Value Sets",x = TaVS[1,], startRow = 1, headerStyle = dd_header_style)

#Adjustments to the notebook to make it easier to read after initial creation.
for (sheet in wb$sheet_names){
  setColWidths(wb = wb,cols = 1:50, sheet = sheet,widths = 25)
  freezePane(wb = wb, sheet = sheet, firstRow = TRUE)
}
setColWidths(wb = wb,cols = 4, sheet = "Terms and Value Sets",widths = 200)


#Write out workbook
saveWorkbook(wb = wb,file = paste(path,output_file,".xlsx",sep = ""), overwrite = TRUE)

cat(paste("\n\nProcess Complete.\n\nThe output file can be found here: ",path,"\n\n",sep = "")) 
