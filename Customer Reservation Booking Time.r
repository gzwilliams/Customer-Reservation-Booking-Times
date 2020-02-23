#### How Far out from the reserved time do customer's make bookings? ####

##FOR SCRIPT TESTING PURPOSES 
client="xxxxxx";facility = " ";display_data=0;year="2019"

 
## bring in relevant packages  
library(ggiraph)
library(ggplot2)
library(lubridate)
library(tidyr)
library(DT)
library(grid)
library(scales)
library(htmlwidgets)
library(xlsx) #puts all info into excel workbook



## Pull in 2019 Data ##
minusdata3 <- read.csv(paste0("AllClientData",year,".csv"),stringsAsFactors = FALSE)

## correct the 1st name column
colnames(minusdata3)[1] <- "Client"

#### FILTER BY Client ####
minusdata3 = minusdata3[which(minusdata3$Client %in% client),]

## to make sure the client doesn't have a subset account
if(facility != " "){
  minusdata3 = minusdata3[which(minusdata3$Facility == facility),]
}
#else minusdata3$Facility = NA #there is no Facility column in minusdata3

## Remove NAs
minusdata3 <- na.omit(minusdata3)

#### IF SQLData SELECTED CLIENT HAS ZERO rows - then put up warning message ####
if(nrow(minusdata3) < 1){
  nullplot = ggplot(minusdata3) + geom_blank()
  return(list(nullplot,nullplot,nullplot,nullplot,nullplot,nullplot,nullplot,nullplot,nullplot, nullplot,
              nullplot, nullplot, nullplot, nullplot, nullplot, nullplot, nullplot, nullplot, nullplot,
              nullplot, nullplot, nullplot, nullplot, nullplot, nullplot, nullplot, nullplot,
              nullplot, nullplot, nullplot, "NOT ENOUGH DATA"))
}

## MySQL Server changed the HoursBookedBeforeTT and BTHour columns into characters
## must change back to integer form
minusdata3$HoursBookedBeforeTT <- as.numeric(minusdata3$HoursBookedBeforeTT, na.rm=TRUE)
minusdata3$BTHour <- as.numeric(minusdata3$BTHour, na.rm=TRUE)

minusdata3 <- na.omit(minusdata3) #omit NAs

#### Remove data points with HoursBookedBeforeTT that is less than -336 hours(further than 2weeks out) ####
minusdata3 = minusdata3[!minusdata3$HoursBookedBeforeTT < -336,]

#### Create background  image w/ opacity from 5-15% ####
image3 <- jpeg::readJPEG("5-15%OpacityImage.jpg")
imagex <- rasterGrob(image3, width = unit(1, "npc"), height = unit(1, "npc"))


##### Convert DayType from factor to character ###### Not needed if stringsasFactors=False


#### Format the BookingTime Column ###
minusdata3$BookingTime <- as.POSIXct(minusdata3$BookingTime, format="%H:%M:%S")

#### Separate by DayType - Weekday.Weekend NEEDED #####
weekday.vec <- (minusdata3[minusdata3$DayType=="Weekday",])
weekend.vec <- (minusdata3[minusdata3$DayType=="Weekend",])

##### Histograms of HoursBookedBeforeTT ####

#### tooltip created ####
tooltip_css <- "background-color:gray;color:white;font-style:italic;padding:10px;border-radius:10px 20px 10px 20px;font-family:Trebuchet MS"

#### Create TOTAL Histogram Y-axis limit for annotations ####
ymaxtenth <- nrow(minusdata3[minusdata3$HoursBookedBeforeTT > median(minusdata3$HoursBookedBeforeTT,na.rm=TRUE),])

#### Interactive Histogram by Channel #### totalmin
totmin <- ggplot(minusdata3, aes(HoursBookedBeforeTT, color = Channel, fill = Channel, tooltip=Channel, data_id=Channel)) + annotation_custom(imagex) + 
  geom_histogram_interactive() +
  theme(text=element_text(family="Trebuchet MS", face="bold")) #+
  
totalmin <- totmin + labs(x = "Hours Before Reservation", y = "Frequency of Bookings", title = paste0(minusdata3$Client, " Total ", minusdata3$Year, " Distribution by Channel"), subtitle=paste0("Median Hours Booked in Advance Before Reservation: ",median(minusdata3$HoursBookedBeforeTT, na.rm=TRUE))) +
  geom_vline_interactive(aes(xintercept = median(minusdata3$HoursBookedBeforeTT, na.rm=TRUE), tooltip=paste0("Median: ", median(minusdata3$HoursBookedBeforeTT, na.rm=TRUE)), data_id=median(minusdata3$HoursBookedBeforeTT, na.rm=TRUE)), color="yellow", size=0.7) +
  geom_vline_interactive(aes(xintercept = -24, tooltip=paste0("1 Day in Advance")), color="green") +
  geom_vline_interactive(aes(xintercept = -48, tooltip=paste0("2 Days in Advance")), color="green") +
  geom_vline_interactive(aes(xintercept = -72, tooltip=paste0("3 Days in Advance")), color="green") +
  geom_vline_interactive(aes(xintercept = -96, tooltip=paste0("4 Days in Advance")), color="green") +
  geom_vline_interactive(aes(xintercept = -120, tooltip=paste0("5 Days in Advance")), color="green") +
  geom_vline_interactive(aes(xintercept = -144, tooltip=paste0("6 Days in Advance")), color="green") +
  geom_vline_interactive(aes(xintercept = -168, tooltip=paste0("7 Days in Advance")), color="green") +
  geom_vline_interactive(aes(xintercept = -192, tooltip=paste0("8 Days in Advance")), color="green") +
  geom_vline_interactive(aes(xintercept = -216, tooltip=paste0("9 Days in Advance")), color="green") +
  geom_vline_interactive(aes(xintercept = -240, tooltip=paste0("10 Days in Advance")), color="green") +
  geom_vline_interactive(aes(xintercept = -264, tooltip=paste0("11 Days in Advance")), color="green") +
  geom_vline_interactive(aes(xintercept = -288, tooltip=paste0("12 Days in Advance")), color="green") +
  geom_vline_interactive(aes(xintercept = -312, tooltip=paste0("13 Days in Advance")), color="green") +
  geom_vline_interactive(aes(xintercept = -336, tooltip=paste0("14 Days in Advance")), color="green") +
  annotate(geom="label",x=-24, y=(ymaxtenth*.95), label="1DayAdvance", color="black", fill="green", family="Trebuchet MS", fontface="bold", size=2.5, vjust=1) +
  annotate(geom="label",x=-48, y=(ymaxtenth*.90), label="2DaysAdvance", color="black", fill="green", family="Trebuchet MS", fontface="bold", size=2.5, vjust=1) +
  annotate(geom="label",x=-72, y=(ymaxtenth*.85), label="3DaysAdvance", color="black", fill="green", family="Trebuchet MS", fontface="bold", size=2.5, vjust=1) +
  annotate(geom="label",x=-96, y=(ymaxtenth*.80), label="4DaysAdvance", color="black", fill="green", family="Trebuchet MS", fontface="bold", size=2.5, vjust=1) +
  annotate(geom="label",x=-120, y=(ymaxtenth*.75), label="5DaysAdvance", color="black", fill="green", family="Trebuchet MS", fontface="bold", size=2.5, vjust=1) +
  annotate(geom="label",x=-144, y=(ymaxtenth*.70), label="6DaysAdvance", color="black", fill="green", family="Trebuchet MS", fontface="bold", size=2.5, vjust=1) +
  annotate(geom="label",x=-168, y=(ymaxtenth*.65), label="7DaysAdvance", color="black", fill="green", family="Trebuchet MS", fontface="bold", size=2.5, vjust=1) +
  annotate(geom="label",x=-192, y=(ymaxtenth*.60), label="8DaysAdvance", color="black", fill="green", family="Trebuchet MS", fontface="bold", size=2.5, vjust=1) +
  annotate(geom="label",x=-216, y=(ymaxtenth*.55), label="9DaysAdvance", color="black", fill="green", family="Trebuchet MS", fontface="bold", size=2.5, vjust=1) +
  annotate(geom="label",x=-240, y=(ymaxtenth*.50), label="10DaysAdvance", color="black", fill="green", family="Trebuchet MS", fontface="bold", size=2.5, vjust=1) +
  annotate(geom="label",x=-264, y=(ymaxtenth*.45), label="11DaysAdvance", color="black", fill="green", family="Trebuchet MS", fontface="bold", size=2.5, vjust=1) +
  annotate(geom="label",x=-288, y=(ymaxtenth*.40), label="12DaysAdvance", color="black", fill="green", family="Trebuchet MS", fontface="bold", size=2.5, vjust=1) +
  annotate(geom="label",x=-312, y=(ymaxtenth*.35), label="13DaysAdvance", color="black", fill="green", family="Trebuchet MS", fontface="bold", size=2.5, vjust=1) +
  annotate(geom="label",x=-336, y=(ymaxtenth*.30), label="14DaysAdvance", color="black", fill="green", family="Trebuchet MS", fontface="bold", size=2.5, vjust=1, hjust=.2) 

#print out in results - old ggiraph output function
#ggiraph( code = print(totalmin), width_svg = 8, height_svg = 4, tooltip_extra_css = tooltip_css, hover_css = "cursor:pointer;fill:red;r:8pt;")

#### Total Histogram Density Channel ####
totalmin.denz <- ggplot(minusdata3, aes(HoursBookedBeforeTT, color = Channel, fill = Channel, tooltip=Channel, data_id=Channel)) +
  annotation_custom(imagex) + 
  theme(text=element_text(family="Trebuchet MS", face="bold")) +
  labs(x = "Hours Before Reservation", y = "Density of Bookings", title = paste0(minusdata3$Client, " Total ", minusdata3$Year, " Density Distribution by Channel"), subtitle=paste0("Median Hours Booked in Advance Before Reservation: ",median(minusdata3$HoursBookedBeforeTT, na.rm=TRUE))) +
  geom_vline_interactive(aes(xintercept = median(minusdata3$HoursBookedBeforeTT, na.rm=TRUE), tooltip=paste0("Median: ", median(minusdata3$HoursBookedBeforeTT, na.rm=TRUE)), data_id=median(minusdata3$HoursBookedBeforeTT, na.rm=TRUE)), color="yellow", size=0.7) +
  geom_density(alpha=0.2)

#ggiraph( code = print(totalmin.denz), width_svg = 8, height_svg = 4, tooltip_extra_css = tooltip_css, hover_css = "cursor:pointer;fill:red;r:8pt;")


#### Total Histogram by Rounds #### totalmin.rnd
totmin.rnd <- ggplot(minusdata.rnd, aes(HoursBookedBeforeTT, color = Rounds, fill = Rounds, tooltip=Rounds, data_id=Rounds)) + annotation_custom(imagex) + 
  geom_histogram_interactive(aes(tooltip=Rounds, data_id=Rounds)) #+ 
  #geom_histogram() +
  #theme(text=element_text(family="Trebuchet MS", face="bold")) +
  #labs(x = "Hours Before Reservation", y = "Bookings: Frequency/Count", color="Golfers Per Booking", fill="Golfers Per Booking", title = paste0(minusdata3$Client, " Total ", minusdata3$Year, " Distribution by Golfers: Hours Booked in Advance"), subtitle=paste0("Median Hours Booked in Advance Before Reservation: ",median(minusdata.rnd$HoursBookedBeforeTT, na.rm=TRUE))) +
totalmin.rnd <- totmin.rnd +
  geom_vline_interactive(aes(xintercept = median(minusdata.rnd$HoursBookedBeforeTT, na.rm=TRUE), tooltip=paste0("Median: ", median(minusdata.rnd$HoursBookedBeforeTT, na.rm=TRUE)), data_id=median(minusdata.rnd$HoursBookedBeforeTT, na.rm=TRUE)), color="yellow", size=0.7) +
  geom_vline_interactive(aes(xintercept = -24, tooltip=paste0("1 Day in Advance")), color="green") +
  geom_vline_interactive(aes(xintercept = -48, tooltip=paste0("2 Days in Advance")), color="green") +
  geom_vline_interactive(aes(xintercept = -72, tooltip=paste0("3 Days in Advance")), color="green") +
  geom_vline_interactive(aes(xintercept = -96, tooltip=paste0("4 Days in Advance")), color="green") +
  geom_vline_interactive(aes(xintercept = -120, tooltip=paste0("5 Days in Advance")), color="green") +
  geom_vline_interactive(aes(xintercept = -144, tooltip=paste0("6 Days in Advance")), color="green") +
  geom_vline_interactive(aes(xintercept = -168, tooltip=paste0("7 Days in Advance")), color="green") +
  geom_vline_interactive(aes(xintercept = -192, tooltip=paste0("8 Days in Advance")), color="green") +
  geom_vline_interactive(aes(xintercept = -216, tooltip=paste0("9 Days in Advance")), color="green") +
  geom_vline_interactive(aes(xintercept = -240, tooltip=paste0("10 Days in Advance")), color="green") +
  geom_vline_interactive(aes(xintercept = -264, tooltip=paste0("11 Days in Advance")), color="green") +
  geom_vline_interactive(aes(xintercept = -288, tooltip=paste0("12 Days in Advance")), color="green") +
  geom_vline_interactive(aes(xintercept = -312, tooltip=paste0("13 Days in Advance")), color="green") +
  geom_vline_interactive(aes(xintercept = -336, tooltip=paste0("14 Days in Advance")), color="green") +
  annotate(geom="label",x=-24, y=(ymaxtenth*.95), label="1DayAdvance", color="black", fill="green", family="Trebuchet MS", fontface="bold", size=2.5, vjust=1) +
  annotate(geom="label",x=-48, y=(ymaxtenth*.90), label="2DaysAdvance", color="black", fill="green", family="Trebuchet MS", fontface="bold", size=2.5, vjust=1) +
  annotate(geom="label",x=-72, y=(ymaxtenth*.85), label="3DaysAdvance", color="black", fill="green", family="Trebuchet MS", fontface="bold", size=2.5, vjust=1) +
  annotate(geom="label",x=-96, y=(ymaxtenth*.80), label="4DaysAdvance", color="black", fill="green", family="Trebuchet MS", fontface="bold", size=2.5, vjust=1) +
  annotate(geom="label",x=-120, y=(ymaxtenth*.75), label="5DaysAdvance", color="black", fill="green", family="Trebuchet MS", fontface="bold", size=2.5, vjust=1) +
  annotate(geom="label",x=-144, y=(ymaxtenth*.70), label="6DaysAdvance", color="black", fill="green", family="Trebuchet MS", fontface="bold", size=2.5, vjust=1) +
  annotate(geom="label",x=-168, y=(ymaxtenth*.65), label="7DaysAdvance", color="black", fill="green", family="Trebuchet MS", fontface="bold", size=2.5, vjust=1) +
  annotate(geom="label",x=-192, y=(ymaxtenth*.60), label="8DaysAdvance", color="black", fill="green", family="Trebuchet MS", fontface="bold", size=2.5, vjust=1) +
  annotate(geom="label",x=-216, y=(ymaxtenth*.55), label="9DaysAdvance", color="black", fill="green", family="Trebuchet MS", fontface="bold", size=2.5, vjust=1) +
  annotate(geom="label",x=-240, y=(ymaxtenth*.50), label="10DaysAdvance", color="black", fill="green", family="Trebuchet MS", fontface="bold", size=2.5, vjust=1) +
  annotate(geom="label",x=-264, y=(ymaxtenth*.45), label="11DaysAdvance", color="black", fill="green", family="Trebuchet MS", fontface="bold", size=2.5, vjust=1) +
  annotate(geom="label",x=-288, y=(ymaxtenth*.40), label="12DaysAdvance", color="black", fill="green", family="Trebuchet MS", fontface="bold", size=2.5, vjust=1) +
  annotate(geom="label",x=-312, y=(ymaxtenth*.35), label="13DaysAdvance", color="black", fill="green", family="Trebuchet MS", fontface="bold", size=2.5, vjust=1) +
  annotate(geom="label",x=-336, y=(ymaxtenth*.30), label="14DaysAdvance", color="black", fill="green", family="Trebuchet MS", fontface="bold", size=2.5, vjust=1, hjust=.2) +
  theme(text=element_text(family="Trebuchet MS", face="bold")) +
  labs(x = "Hours Before Reservation", y = "Frequency of Bookings", color="Golfers Per Booking", fill="Golfers Per Booking", title = paste0(minusdata3$Client, " Total ", minusdata3$Year, " Distribution by Golfers"), subtitle=paste0("Median Hours Booked in Advance Before Reservation: ",median(minusdata.rnd$HoursBookedBeforeTT, na.rm=TRUE))) 

#ggiraph( code = print(totalmin.rnd), width_svg = 8, height_svg = 4, tooltip_extra_css = tooltip_css, hover_css = "cursor:pointer;fill:red;r:1pt;")
#ggiraph( code = print(totmin.rnd), width_svg = 8, height_svg = 4, tooltip_extra_css = tooltip_css, hover_css = "cursor:pointer;fill:red;r:1pt;")

#### Total Histogram Density Rounds ####
totalmin.rnd.denz <- ggplot(minusdata.rnd, aes(HoursBookedBeforeTT, color = Rounds, fill = Rounds, tooltip=Rounds, data_id=Rounds)) +
  annotation_custom(imagex) + 
  theme(text=element_text(family="Trebuchet MS", face="bold")) +
  labs(x = "Hours Before Reservation", y = "Density of Bookings", color="Golfers Per Booking", fill="Golfers Per Booking", title = paste0(minusdata.rnd$Client, " Total ", minusdata.rnd$Year, " Density Dist. by Golfers"), subtitle=paste0("Median Hours Booked in Advance Before Reservation: ",median(minusdata.rnd$HoursBookedBeforeTT, na.rm=TRUE))) +
  geom_vline_interactive(aes(xintercept = median(minusdata.rnd$HoursBookedBeforeTT, na.rm=TRUE), tooltip=paste0("Median: ", median(minusdata.rnd$HoursBookedBeforeTT, na.rm=TRUE)), data_id=median(minusdata.rnd$HoursBookedBeforeTT, na.rm=TRUE)), color="yellow", size=0.7) +
  geom_density(alpha=0.2)

#ggiraph( code = print(totalmin.rnd.denz), width_svg = 8, height_svg = 4, tooltip_extra_css = tooltip_css, hover_css = "cursor:pointer;fill:red;r:8pt;")



#### Booking Time Graphs ####

## Format the BookingTime Column
minusdata3$BookingTime <- as.POSIXct(minusdata3$BookingTime, format="%H:%M:%S")

#### Total Booking Time by Channel - COUNT ####
totalbooktime.chan <- ggplot(minusdata3, aes(BookingTime, stat(count), color = Channel, fill = Channel, tooltip=Channel, data_id=Channel)) + annotation_custom(imagex) + 
  geom_histogram_interactive(bins=48) + 
  theme(text=element_text(family="Trebuchet MS", face="bold"), axis.text.x = element_text(vjust=.5, size=8)) +
  labs(x = "Time of Day", y = "Frequency/Count of Bookings", title = paste0(minusdata3$Client," - Total ",minusdata3$Year, " Distribution: Booking Time by Channel"), subtitle=paste0("Median Booking Time: ",format(as.POSIXct(median(minusdata3$BookingTime, na.rm=TRUE),format="%H:%M"),"%I:%M %p"))) +
  scale_x_datetime(date_breaks = "2 hour", date_labels = "%l%p")

#ggiraph( code = print(totalbooktime.chan), width_svg = 8, height_svg = 4, tooltip_extra_css = tooltip_css, hover_css = "cursor:pointer;fill:chartreuse;r:2pt;")


#### TOTAL Booking Time Density Charts by Channel - DENSITY CHANNEL ####
totalbktme.denz.chan <- ggplot(minusdata3, aes(BookingTime, color=Channel, fill=Channel)) + 
  annotation_custom(imagex) +
  labs(x = "Time of Day", y = "Density of Bookings", title = paste0(minusdata.rnd$Client,": Total ",minusdata.rnd$Year, " Booking Time Density Distribution"), subtitle=paste0("Booking Time Density by Channel Per Booking")) +
  theme(text=element_text(family="Trebuchet MS", face="bold"), axis.text.x = element_text(vjust=.5, size=8)) +
  scale_x_datetime(date_breaks = "2 hour", date_labels = "%l%p") +
  geom_density(alpha=0.2) 

#ggiraph( code = print(totalbktme.denz.chan), width_svg = 8, height_svg = 4, tooltip_extra_css = tooltip_css, hover_css = "cursor:pointer;fill:red;r:2pt;")


#### ggiraph outputs ####

HBBTT.tot.chan <- ggiraph(code = print(totalmin), width_svg = 8, height_svg = 4, selection_type ="none", tooltip_extra_css = tooltip_css, hover_css = "cursor:pointer;fill:red;r:1pt;")
HBBTT.tot.chan.dnz <- ggiraph(code = print(totalmin.denz), width_svg = 8, height_svg = 4, selection_type ="none", tooltip_extra_css = tooltip_css, hover_css = "cursor:pointer;fill:red;r:1pt;")

BKT.tot.chan <- ggiraph( code = print(totalbooktime.chan), width_svg = 8, height_svg = 4, selection_type ="none", tooltip_extra_css = tooltip_css, hover_css = "cursor:pointer;fill:red;r:1pt;")
BKT.tot.chan.dnz <- ggiraph( code = print(totalbktme.denz.chan), width_svg = 8, height_svg = 4, selection_type ="none", tooltip_extra_css = tooltip_css, hover_css = "cursor:pointer;fill:red;r:1pt;")


#### DATA TABLES for Channels-Hours-Median HoursBookedBeforeTT #####

####TOTAL TABLE - Median Hours by Channel/Hour - CHANNEL ####
finaltot.chan.tbl <- with(minusdata3, tapply(HoursBookedBeforeTT, list(TTHour, Channel), median))
## make sure NAs are seen as NAs/blanks
finaltot.chan.tbl[which(is.na(finaltot.chan.tbl),arr.ind = TRUE)] <- NA
## use cbind to assign column title to rownames column
finaltot.chan.tbl <- as.data.frame(finaltot.chan.tbl)
finaltot.chan.tbl <- cbind(Hour = rownames(finaltot.chan.tbl), finaltot.chan.tbl)
#Create the final table with formatting, buttons, DL file format, etc.
fin.tot.chan.tbl <- datatable(finaltot.chan.tbl, rownames = FALSE,
                              caption = htmltools::tags$caption(paste0(minusdata3$Client,' - Total ',minusdata3$Year,' ','Table: Median Hours Booked in Advanced-Channel')),
                              extensions = 'Buttons', options=list(
                                dom= 'Bfrtip',
                                buttons=list('copy',list(extend='excel',filename=paste0(minusdata3$Client,' - Total ', minusdata3$Year,' Table- Median Hrs Booked Adv - Channel',Sys.Date()))),
                                pageLength=nrow(finaltot.chan.tbl)
                              ))



#### BOOKING TIME HEAT MAP ####
BKTHeat <-read.csv(paste0("BookingTimeHeatMap",year,".csv"),stringsAsFactors = FALSE)

#make sure the Client column is named properly
colnames(BKTHeat)[1] <- "Client"

### FILTER BY Client ###
BKTHeat = BKTHeat[which(BKTHeat$Client %in% Client),]

## Remove NAs ##
BKTHeat <- na.omit(BKTHeat)

##create columns for diff booking times##
BKTHeat$BTHourlp = as.POSIXct(as.POSIXct(strptime(BKTHeat$BTHour, "%H")),"%l%p")
BKTHeat$BTHour = as.POSIXct(strptime(BKTHeat$BTHour, "%H"))

### Order the Days of the week - Ordered.Factor ##
BKTHeat$DayofWeek <- ordered(BKTHeat$DayofWeek, levels=rev(c("Monday", "Tuesday", "Wednesday",
                                                             "Thursday", "Friday", "Saturday",
                                                             "Sunday")))

#### Floating tooltip ####
BKTHeat$tooltipbookings = paste0("Bookings: ", formatC(BKTHeat$Bookings, format="d", big.mark=","), "<br/>",
                                 "Day: ", BKTHeat$DayofWeek, "<br/>", "Hour: ", format(BKTHeat$BTHourlp, format="%l%p"))

##create ggplot for heatmap - count of bookings
htmpg <- ggplot(BKTHeat, aes(BTHour,DayofWeek, fill=Bookings, tooltip=tooltipbookings)) +
  #geom_tile(color="white", size=0.1) +
  geom_tile_interactive(color="white", size=0.1) +
  scale_fill_gradient(low="green", high="red") +
  #scale_fill_gradient2_interactive(low="green", mid="yellow", high="red")
  labs(x = "Time of Day", y = "Day of Week", title = paste0(BKTHeat$Client,": ",BKTHeat$Year, " Booking Time Heat Map")) +
  theme(text=element_text(family="Trebuchet MS", face="bold"), axis.text.x = element_text(vjust=.5, size=8)) +
  scale_x_datetime(date_breaks = "2 hour", date_labels = "%l%p")



#HEATBKmap <- ggiraph(code = print(htmpg), width_svg = 8, height_svg = 4, tooltip_extra_css = tooltip_css)
#HEATBKmap <- girafe(ggobj = htmpg)


#### Frequency Tables next to Density Curves in download workbook####

## manual list of vectors (48) for each permutation of TotWkdyWknd(3) x ChanRnd(2) x
## (ResShopTPYWeb(4) x Golfers1234(4)) = 48 tables
##

#### Hours Booked Before TT vectors ####
totalhb4t.ch.res <- (minusdata3[minusdata3$Channel=="Res Center",]["HoursBookedBeforeTT"])
totalhb4t.ch.shop <- (minusdata3[minusdata3$Channel=="Shop",]["HoursBookedBeforeTT"])
totalhb4t.ch.tpy <- (minusdata3[minusdata3$Channel=="TPY",]["HoursBookedBeforeTT"])
totalhb4t.ch.web <- (minusdata3[minusdata3$Channel=="Website",]["HoursBookedBeforeTT"])


#### Booking Time vectors ####
totalbktd.ch.res <- (minusdata3[minusdata3$Channel=="Res Center",]["BTHour"])
totalbktd.ch.shop <- (minusdata3[minusdata3$Channel=="Shop",]["BTHour"])
totalbktd.ch.tpy <- (minusdata3[minusdata3$Channel=="TPY",]["BTHour"])
totalbktd.ch.web <- (minusdata3[minusdata3$Channel=="Website",]["BTHour"])



## put vectors into a list in order to apply a function
listx <- as.list(c(totalhb4t.ch.res, totalhb4t.ch.shop, totalhb4t.ch.tpy, totalhb4t.ch.web, 
                    totalbktd.ch.res, totalbktd.ch.shop, totalbktd.ch.tpy, totalbktd.ch.web))

listxj <- list(totalhb4t.ch.res, totalhb4t.ch.shop, totalhb4t.ch.tpy, totalhb4t.ch.web, 
               totalbktd.ch.res, totalbktd.ch.shop, totalbktd.ch.tpy, totalbktd.ch.web, na.rm=TRUE)



## possible to name each vector in a list?
## must refer to each data frame as a number in a list as according to the listx(c(.,.,.,)) order above

## Denz8plots function to create frequency, cumualative frequency, relative frequency,
## & cumulative relative frequency table
Denz8Plots <- function(x) {
  
  x <- as.data.frame(x)

  if(colnames(x[1])=="BTHour"){
    testv1 <- factor(cut(x$BTHour, breaks = seq(0, 24, by=3)))
  }
  if(colnames(x[1])=="HoursBookedBeforeTT") {
    #x <- as.data.frame(totalhb4t.ch.res) #used for testing
    testv1 <- factor(cut(x$HoursBookedBeforeTT, breaks = seq(0, -336, by=-24)))
    ##re-order levels to start from 0
    testv1 <- factor(testv1, levels = rev(levels(testv1)))
  }

  testtb <- cbind(Frequency=table(testv1),
                  Cumulative=cumsum(table(testv1)),
                  RelativeFreq=prop.table(table(testv1)),
                  CumulativeRelativeFreq=cumsum((prop.table(table(testv1)))))
  testtb <- as.data.frame(testtb)
  testtb <- cbind(Hours=rownames(testtb), testtb)
  return(testtb)
}

#### Apply the Denz8Plots function onto the list of vectors ####
## now we will be able to see the count and percentages - maybe 80% of bookings are made within 48 hours,
## or maybe 60% of bookings are made more than 7 days in advance?
DenzListPlots <- lapply(seq_along(listx), function(i)Denz8Plots(listx[i]))


#### FULL WORKBOOK DOWNLOAD #### --- Using xlsx package
#Client <- minusdata3$Client[1]
#year <- minusdata3$Year[1]
## change working directory to save image files ##
setwd("/xxx/xx/Client Stuff/Downloads/")
####CREATE WB ####

BkAdvWB <- createWorkbook(type="xlsx")

####Add graphs and tables to 'Hours Booked in Advance-Channel' Sheet ####
Sheetname1 = "Hours Booked in Advance-Channel"
## try creating png image immediatly before dev.off

png(filename = "totalhb4t.chan.count.png", width=800, height=400)
print(totalmin) #total hours before Reservation count graph
dev.off()

png(filename = "totalhb4t.chan.density.png", width=800, height=400)
print(totalmin.denz) #total hours before Reservation density graph
dev.off()

png(filename = "bookingtime.heatmap.png", width=800, height=400)
print(htmpg) #Booking Time Density Chart
dev.off()

#### Create Headers for FreqRelativeCumulative Tables ####
titleResCenter <- paste0("Res Center")
titleShop <- paste0("Shop")
titleTPY <- paste0("TPY")
titleWebsite <- paste0("Website")

#### Create sheet in workbook ####
sheet1 <- createSheet(BkAdvWB, sheetName = Sheetname1)

##create a CellStyle that goes into BkAdvWB
cstylecols <- CellStyle(BkAdvWB) + Font(BkAdvWB, isBold=TRUE) + Border(color="black") 

#Create table headers for Median Hours in Advance #
TotalChanTitle <- paste0(year," Total: Median Hours Booked in Advance by Channel")

#### Creating/Defining the Cells for xlsx text input to WB ####
## xlsx package - getRows from sheet, getCells from rows###
## in xlsx you have to define the cells (rows & cols) for each sheet
rows1 <- createRow(sheet1, 1:120) #creates 120 rows - more than needed in this example
cells1 <- createCell(rows1, colIndex = 1:43) #creates 43 columns & adds them to the defined rows

## 'year' Total: Median Hours Booked in Advance by Channel
##format title cells
cell <- cells1[[21,1]]
setCellValue(cell, TotalChanTitle)
setCellStyle(cell, cstylecols)


#### Freq,Relative,Cum Freq Titles - Res Center ####
cell <- cells1[[1,31]]
setCellValue(cell, titleResCenter)
setCellStyle(cell, cstylecols)


#### Freq,Relative,Cum Freq Titles - TPY ####
cell <- cells1[[18,31]]
setCellValue(cell, titleTPY)
setCellStyle(cell, cstylecols)


#### Freq,Relative,Cum Freq Titles - Shop ####
cell <- cells1[[1,37]]
setCellValue(cell, titleShop)
setCellStyle(cell, cstylecols)


#### Freq,Relative,Cum Freq Titles - Website ####
cell <- cells1[[18,37]]
setCellValue(cell, titleWebsite)
setCellStyle(cell, cstylecols)


#### Channel Tables - Median Hours in Adv ####
addDataFrame(finaltot.chan.tbl, sheet1, startRow=22, startColumn=1, row.names=FALSE, colnamesStyle=cstylecols)

#### Booked in Advance Channel CHARTS ####
addPicture(file=paste0("totalhb4t.chan.count.png"), sheet1, startRow=1, startColumn=1)
addPicture(file=paste0("totalhb4t.chan.density.png"), sheet1, startRow=1, startColumn=16)

#### total channels - Freq,Relative,Cum Freq DATA TABLES  ####
addDataFrame(DenzListPlots[1], sheet1, startRow=2, startColumn=31, row.names=FALSE, colnamesStyle=cstylecols)
addDataFrame(DenzListPlots[2], sheet1, startRow=2, startColumn=37, row.names=FALSE, colnamesStyle=cstylecols)
addDataFrame(DenzListPlots[3], sheet1, startRow=19, startColumn=31, row.names=FALSE, colnamesStyle=cstylecols)
addDataFrame(DenzListPlots[4], sheet1, startRow=19, startColumn=37, row.names=FALSE, colnamesStyle=cstylecols)

#### REMOVED 2nd Sheet ####


#### SHEET 3 'Booking Time Distribution - Chan' #### 
####Add graphs and tables to 'Booking Time Distribution-Channel' Sheet ####
Sheetname3 = "Booking Time Dist.-Channel"

#### Create sheet in workbook ####
sheet3 <- createSheet(BkAdvWB, sheetName = Sheetname3)

#create a CellStyle that goes into BkAdvWB
cstylecols <- CellStyle(BkAdvWB) + Font(BkAdvWB, isBold=TRUE) + Border(color="black")

#### Creating/Defining the Cells for xlsx text input to WB ####
## xlsx package - getRows from sheet, getCells from rows###
## in xlsx you have to define the cells (rows & cols) for each sheet
rows1 <- createRow(sheet3, 1:120) #creates 120 rows
cells1 <- createCell(rows1, colIndex = 1:43) #creates 43 columns & adds them to rows

## try creating png image immediatly before dev.off
png(filename = "totalbktd.chan.count.png", width=800, height=400)
print(totalbooktime.chan) #total hours before Reservation count graph
dev.off()

png(filename = "totalbktd.chan.density.png", width=800, height=400)
print(totalbktme.denz.chan) #total hours before Reservation density graph
dev.off()



#### Freq,Relative,Cum Freq Titles - Res Center ####
cell <- cells1[[1,31]]
setCellValue(cell, titleResCenter)
setCellStyle(cell, cstylecols)

#### Freq,Relative,Cum Freq Titles - TPY ####
cell <- cells1[[12,31]]
setCellValue(cell, titleTPY)
setCellStyle(cell, cstylecols)


#### Freq,Relative,Cum Freq Titles - Shop ####
cell <- cells1[[1,37]]
setCellValue(cell, titleShop)
setCellStyle(cell, cstylecols)


#### Freq,Relative,Cum Freq Titles - Website ####
cell <- cells1[[12,37]]
setCellValue(cell, titleWebsite)
setCellStyle(cell, cstylecols)

#### Booking Time Distribution - Channel CHARTS ####
addPicture(file=paste0("totalbktd.chan.count.png"), sheet3, startRow=1, startColumn=1)
addPicture(file=paste0("totalbktd.chan.density.png"), sheet3, startRow=1, startColumn=16)


#### total channel booktime dist - Freq,Relative,Cum Freq DATA TABLES  ####
addDataFrame(DenzListPlots[25], sheet3, startRow=2, startColumn=31, row.names=FALSE, colnamesStyle=cstylecols)
addDataFrame(DenzListPlots[26], sheet3, startRow=2, startColumn=37, row.names=FALSE, colnamesStyle=cstylecols)
addDataFrame(DenzListPlots[27], sheet3, startRow=13, startColumn=31, row.names=FALSE, colnamesStyle=cstylecols)
addDataFrame(DenzListPlots[28], sheet3, startRow=13, startColumn=37, row.names=FALSE, colnamesStyle=cstylecols)


#### REMOVED SHEET 4  ####

#### SHEET 5 'Booking Time HeatMap' ####
Sheetname5 = "Booking Time HeatMap"

#### Create sheet in workbook ####
sheet5 <- createSheet(BkAdvWB, sheetName = Sheetname5)

addPicture(file=paste0("bookingtime.heatmap.png"), sheet5, startRow=1, startColumn=1)


 
saveWorkbook(BkAdvWB)
