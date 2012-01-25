#!/usr/bin/env python

# The script creates a set of Cards Against Humanity from a CSV file.
# http://cardsagainsthumanity.com/
#
# Scribus script > 1.4 compatible
# Based on uno.py by Giorgos Logiotatidis and importcvs2table by Sebastian Stetter
# Lang HOANG (hoanglang@gmail.com)
# Under the GPL Licence


# Usage : 
# - launch script from Scribus
# - Enter the CSV infos (delimiter char, quote and column to process)
# - Select the CSV file
#
# The default format is A4 - landscape.

from __future__ import division
import sys

try:
    # Please do not use 'from scribus import *' . If you must use a 'from import',
    # Do so _after_ the 'import scribus' and only import the names you need, such
    # as commonly used constants.
    import scribus
except ImportError,err:
    print "This Python script is written for the Scribus scripting interface."
    print "It can only be run from within Scribus."
    sys.exit(1)

#########################
# YOUR IMPORTS GO HERE  #
#########################
import csv

# Constants
WIDTH=297 # page width
HEIGHT=210 # page height
MARGINS=(10, 10, 10, 10) # margins (left, right, top, bottom) ?
CARDWIDTH=69 # card width
CARDHEIGHT=47.5 # card height

# Get the cvs data
def getCSVdata(delim, qc):
    """opens a csv file, reads it in and returns a 2 dimensional list with the data"""
    csvfile = scribus.fileDialog("csv2table new:: open file", "*.csv")
    if csvfile != "":
        try:
            reader = csv.reader(file(csvfile), delimiter=delim, quotechar=qc)
            datalist=[]
            skipfirst=False
            for row in reader:
                if skipfirst==True:
                    rowlist=[]
                    for col in row:
                        rowlist.append(col)
                    datalist.append(rowlist)
                else : skipfirst=True
            return datalist
        except Exception,  e:
            scribus.messageBox("csv2table new", "Could not open file %s"%e)
    else:
        sys.exit

# compute the number of cards per line/row
def get_multipl(total, width, marginl=0, marginr=0):
    e=divmod(total-(marginl+marginr), width)
    return e[0]

# add the CAH logo
def addLogo(x, y, isBlack):
    #square side
    a=5
    #line
    line=0.01
    #angle
    angle=15
    #LightGrey
    scribus.defineColor("LightGrey",0,0,0,128)
    #DarkGrey
    scribus.defineColor("DarkGrey",0,0,0,192)
    
    # White cards colors
    squareColor1="Black"
    lineColor1="White"
    squareColor2="DarkGrey"
    lineColor2="White"
    squareColor3="LightGrey"
    lineColor3="White"
    titleColor="Black"
    
    # Black cards colors
    if isBlack==True:
        squareColor1="DarkGrey"
        lineColor1="DarkGrey"
        squareColor2="LightGrey"
        lineColor2="LightGrey"
        squareColor3="White"
        lineColor3="White"
        titleColor="White"
    
    square1=scribus.createRect(x, y+a/10, a, a)
    scribus.setLineWidth(1, square1)
    scribus.setFillColor(squareColor1, square1)
    scribus.setLineColor(lineColor1, square1)
    scribus.rotateObject(angle, square1)
    
    square2=scribus.createRect(x+a/2, y, a, a)
    scribus.setLineWidth(1, square2)
    scribus.setFillColor(squareColor2, square2)
    scribus.setLineColor(lineColor2, square2)
    
    square3=scribus.createRect(x+a, y, a, a)
    scribus.setLineWidth(1, square3)
    scribus.setFillColor(squareColor3, square3)
    scribus.setLineColor(lineColor3, square3)
    scribus.rotateObject(-angle, square3)
    
    title=scribus.createText(x+10,y+2, 50, 5)
    scribus.setFont("Arial Bold", title)
    scribus.setFontSize(8, title)
    scribus.insertText("Cards Against Humanity", 0, title)
    scribus.setTextColor(titleColor, title)
    return


# Create a new card
# Order of the layers :
# 1-Box
# 2-TextBox
# 3-Logo
def createCell(text, roffset, coffset, w, h, marginl, margint, isBlack):
    #Create textbox
    box=scribus.createRect(marginl+(roffset-1)*w,margint+(coffset-1)*h, w, h)
    textBox=scribus.createText(marginl+(roffset-1)*w,margint+(coffset-1)*h, w, h)
    #insert the text into textbox
    scribus.setFont("Arial Bold", textBox)
    
    
    
    #Default  font size
    fontSize=12
    if len(text) < 50: fontSize=14
    if len(text) < 20: fontSize=16
    scribus.setFontSize(fontSize, textBox)
    scribus.setTextDistances(3, 3, 3, 3, textBox)
    scribus.insertText(text, 0, textBox)

    #black card
    if isBlack==True:
        scribus.setFillColor("Black", box)
        scribus.setTextColor("White", textBox)
        scribus.setLineColor("White", box)


    #add Logo
    hOffset=37
    wOffset=5
    addLogo(marginl+(roffset-1)*w+wOffset,margint+(coffset-1)*h+hOffset, isBlack)
    return


# Main
colstotal=get_multipl(WIDTH, CARDWIDTH, MARGINS[0], MARGINS[1])
rowstotal=get_multipl(HEIGHT, CARDHEIGHT, MARGINS[2], MARGINS[3])
#print "n per rows: "+ str(colstotal) + ", n per cols: " + str(rowstotal) 

# create a new document 
t=scribus.newDocument((WIDTH,HEIGHT), MARGINS, scribus.PORTRAIT, 1, scribus.UNIT_MILLIMETERS, scribus.FACINGPAGES, scribus.FIRSTPAGERIGHT,1)
cr=1
cc=1
nol=0

# ask for CSV infos
tstr=scribus.valueDialog('Cvs Delimiter, Quote and Column to process','Type 1 delimiter, 1 quote character and the column number (Clear for default ,"0):',';"0')
if len(tstr) > 0: delim=tstr[0] 
else: delim=','
if len(tstr) > 1: qc=tstr[1]
else: qc='"'
if len(tstr) > 2: numcol=int(tstr[2])
else: numcol=0

# select black or white cards
color=scribus.valueDialog('Color of the cards :','black (b) or white (w)','w')
if len(color) > 0 and 'b'==color[0]: isBlack=True
else: isBlack=False


# open CSV file
data = getCSVdata(delim=delim, qc=qc)

# Process data
scribus.messagebarText("Processing "+str(nol)+" elements")
scribus.progressTotal(len(data))
for row in data:
    scribus.messagebarText("Processing "+str(nol)+" elements")
    celltext = row[numcol].strip()
    if len(celltext)!=0:
        createCell(celltext, cr, cc, CARDWIDTH, CARDHEIGHT, MARGINS[0], MARGINS[2], isBlack)
        nol=nol+1
        if cr==colstotal and cc==rowstotal:
            #create new page
            scribus.newPage(-1)
            scribus.gotoPage(scribus.pageCount())
            cr=1
            cc=1
        else:
            if cr==colstotal:
                cr=1
                cc=cc+1
            else:
                cr=cr+1
        scribus.progressSet(nol)
scribus.messagebarText("Processed "+str(nol)+" items. ")
scribus.progressReset()

