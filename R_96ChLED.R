# Overview --------------------
# The script loads the excel file specified on line BLA with the sceme for the cycles as dInput.
# Then it resolves the cycles into single commands. So instead having one row with a command 
# which should be cycled 10 times, it makes 20 rows with time when the LED (or group of LEDs) 
# should be switched on or off. 
# 
# After that it creates a new column with the Serial Commands for each row. 
# 
# At least it runs trough the table and sends each serial commands. Hereby the script breaks, 
# so the commands are executed at the right time. 

# Load nessesary libarys
library(dplyr) # universal manipulation of tables
library(stringr)
library(tidyr) 
library(serial) # to send the commands over RS-232 port
library(readxl) # to read the input excel file

# Specify the input exel file here ("./" means it the same directory)
input_file <- "./96ChLED.xlsx"

# # Uncomment this line to enable the file picker function (useful if input files always changes)
# input_file <- file.choose()



# Import the excel file as dInput (dInput) -----------------------
# dInput (dataInput) is the Input with the Groups convertet to wells. 
# Additionally the Time is convertet to Time_ms and the Duration is the whole Time of one cycle in ms. 

dInput <-  read_xlsx(input_file, sheet = "Program") %>%  
  mutate( # Replace Row and Columns shortcuts with the actuall well number 
    Group = Wells, 
    Wells = gsub("R1", "1,9,17,25,33,41,49,57,65,73,81,89", Wells),
    Wells = gsub("R2", "2,10,18,26,34,42,50,58,66,74,82,90", Wells),
    Wells = gsub("R3", "3,11,19,27,35,43,51,59,67,75,83,91", Wells),
    Wells = gsub("R4", "4,12,20,28,36,44,52,60,68,76,84,92", Wells),
    Wells = gsub("R5", "5,13,21,29,37,45,53,61,69,77,85,93", Wells),
    Wells = gsub("R6", "6,14,22,30,38,46,54,62,70,78,86,94", Wells),
    Wells = gsub("R7", "7,15,23,31,39,47,55,63,71,79,87,95", Wells),
    Wells = gsub("R8", "8,16,24,32,40,48,56,64,72,80,88,96", Wells),
    Wells = gsub("C10", "73,74,75,76,77,78,79,80", Wells),
    Wells = gsub("C11", "81,82,83,84,85,86,87,88", Wells),
    Wells = gsub("C12", "89,90,91,92,93,94,95,96", Wells),
    Wells = gsub("C1", "1,2,3,4,5,6,7,8", Wells),
    Wells = gsub("C2", "9,10,11,12,13,14,15,16", Wells),
    Wells = gsub("C3", "17,18,19,20,21,22,23,24", Wells),
    Wells = gsub("C4", "25,26,27,28,29,30,31,32", Wells),
    Wells = gsub("C5", "33,34,35,36,37,38,39,40", Wells),
    Wells = gsub("C6", "41,42,43,44,45,46,47,48", Wells),
    Wells = gsub("C7", "49,50,51,52,53,54,55,56", Wells),
    Wells = gsub("C8", "57,58,59,60,61,62,63,64", Wells),
    Wells = gsub("C9", "65,66,67,68,69,70,71,72", Wells),
    Wells = gsub("G1", "GL 1", Wells),
    Wells = gsub("G2", "GL 2", Wells),
    Wells = gsub("G3", "GL 3", Wells),
    Wells = gsub("G4", "GL 4", Wells),
    Wells = gsub("G5", "GL 5", Wells),
    Wells = gsub("G6", "GL 6", Wells),
    Wells = gsub("G7", "GL 7", Wells),
    Wells = gsub("G8", "GL 8", Wells),
    Wells = gsub("All", "AL", Wells)
  ) %>%
  mutate(Delay_ms = Delay*1000,  # Convert all times in milliseconds
         Duration_ms = ifelse(u_Cycle == "Sec", t_Cycle*1000, 0)+ifelse(u_Cycle == "Min", t_Cycle*60000,0)+ifelse(u_Cycle == "Millisec", t_Cycle,0),
         Time_ms = ifelse(u_Time == "Sec", Time*1000, 0)+ifelse(u_Time == "Min", Time*60000,0)+ifelse(u_Time == "Millisec", Time,0)) %>%
  select(Group, Wells, Cylcle, Delay_ms, Duration_ms, Time_ms, Intensity, u_Time, u_Cycle) %>%
  tbl_df()


# Check if all Unites are known (no spelling mistakes)
unites <- c(dInput$u_Cycle, dInput$u_Time) %>% unique()

if(prod(unites == "Sec" | unites =="Millisec" | unites == "Min") == 0){
  stop(paste("ERROR: Unite(s)", unites ,  "not known!\r"))
}
remove(unites)


# Make d (main commands with Start_Time and End times) ----------------------------------------------------------- 
# in d (data) the cycles od dInput are convertet into single commands, further for every command a start and the end time is
# calculated. 

# create empty dataframe to fill in loop
d <- data.frame(Group = character(), 
                Wells = character(), 
                Start_Time = integer(), 
                End_Time = integer(), 
                Intensity = integer()) %>% 
  tbl_df()


for(Well in dInput$Wells){
  dt <- filter(dInput, Wells == Well)
  
  if(nrow(dt) >1){
    stop("Warning: There are duplicate Group in the CSV file!")
  }
  
  for(i in 1:dt$Cylcle){
    values <- data.frame(Group = dt$Group,
                         Wells = dt$Wells,
                         Start_Time = dt$Duration_ms*(i-1)+dt$Delay_ms,
                         End_Time = dt$Duration_ms*(i-1)+dt$Delay_ms + dt$Time_ms,
                         Intensity = dt$Intensity)
    d <- rbind(d, values)
  }
}
remove(dt, values, i, Well)

d <- d %>% tbl_df() %>% gather(key = Action, value = Time, Start_Time, End_Time) %>% 
  mutate(Action = str_extract(Action, ".*(?=_)")) %>%
  arrange(Time) # sort according to time


## Translate information in serial commands.
createCommand <- function(Wells, Intensity, Action){
  if(Action == "Start"){
    if(grepl("[AG]", Wells)){
      Command <- paste(str_split(Wells, ",")[[1]], " ", Intensity, "\r\n;", sep = "", collapse = "")
    } else {
      Command <- paste("CL ", str_split(Wells, ",")[[1]], " ", Intensity, "\r\n;", sep = "", collapse = "")
    }
    }
  
  if(Action == "End"){
    if(grepl("[AG]", Wells)){
      Command <- paste(str_split(Wells, ",")[[1]], " ", 0, "\r\n;", sep = "", collapse = "")
    } else {
      Command <- paste("CL ", str_split(Wells, ",")[[1]], " ", 0, "\r\n;", sep = "", collapse = "")
    }    }
  return(Command)
}

d <- cbind(d, Command = mapply(createCommand, d$Wells, d$Intensity, d$Action)) %>% tbl_df()



# Start the program / sent the serial commands to the device.-------
con <- serialConnection("COM1") # Start and open connection.
open(con)


start_time <- Sys.time() # Save Program start time. 
print(paste("Run started at: ", Sys.time()))
print(paste("Run will end at:", Sys.time() + d$Time[nrow(d)]/ 1000))
print(paste("Duration:", round(d$Time[nrow(d)]/ 60000, 1),"min."))
print(paste("Duration:", round(d$Time[nrow(d)]/ 3600000,1) , "hour."))


for(i in 1:nrow(d)){
  #while(Sys.time() < df$Time[i] + start_time){} # Wait until the System Time is larger than start time plus the time when the command shall be executed.
  wait <- d$Time[i]-as.numeric(difftime(Sys.time(),start_time, units = "secs")*1000)

  if(wait < 0){
    warning(paste("Group", d$Group[i], "command was at step", i," switched", ifelse(d$Action[i] == "Start", "ON", "OFF"),round(wait*-1), "ms to late."))
    wait <- 0
  }
  Sys.sleep(wait/1000)

  for(j in str_split(str_sub(d$Command[i], end =-2), ";")[[1]]){
    write.serialConnection(con, j)
    Sys.sleep(0.04) # wait 40 ms, else the TouchDevice will miss some commands. 
  }

}

Sys.sleep(0.2)
write.serialConnection(con,"aOFF\r\n")

close(con)
