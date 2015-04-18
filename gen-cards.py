#!/usr/bin/env python3

# Copyright (c) 2015, Stephen Warren. All rights reserved.
#
# Permission is hereby granted, free of charge, to any person obtaining a
# copy of this software and associated documentation files (the "Software"),
# to deal in the Software without restriction, including without limitation
# the rights to use, copy, modify, merge, publish, distribute, sublicense,
# and/or sell copies of the Software, and to permit persons to whom the
# Software is furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in
# all copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.  IN NO EVENT SHALL
# THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
# FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
# DEALINGS IN THE SOFTWARE.

# bootstrap uno component context
import os
import random
import re
import uno
import unohelper

def createUnoService(cClass):
    oServiceManager = smgr
    oObj = oServiceManager.createInstance(cClass)

    return oObj

oCoreReflection = False

def getCoreReflection():
    global oCoreReflection

    if not oCoreReflection:
        oCoreReflection = createUnoService( "com.sun.star.reflection.CoreReflection" )

    return oCoreReflection

def createUnoStruct(cTypeName):
    oCoreReflection = getCoreReflection()

    # Get the IDL class for the type name
    oXIdlClass = oCoreReflection.forName(cTypeName)

    # Create the struct.
    oReturnValue, oStruct = oXIdlClass.createObject(None)

    return oStruct

def createPropertyValue(cName=None, uValue=None, nHandle=None, nState=None):
    oPropertyValue = createUnoStruct("com.sun.star.beans.PropertyValue")

    if cName != None:
        oPropertyValue.Name = cName
    if uValue != None:
        oPropertyValue.Value = uValue
    if nHandle != None:
        oPropertyValue.Handle = nHandle
    if nState != None:
        oPropertyValue.State = nState

    return oPropertyValue

# Get the uno component context from the PyUNO runtime
localContext = uno.getComponentContext()

# Create the UnoUrlResolver
resolver = localContext.ServiceManager.createInstanceWithContext(
    "com.sun.star.bridge.UnoUrlResolver",
    localContext
)

# Connect to the running office
ctx = resolver.resolve(
    "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext"
)

# Retrieve service manager from connection context
smgr = ctx.ServiceManager

# Get the central desktop object
desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

url_base = "file:///" + re.sub(r"\\", "/", os.getcwd()) + "/"

cards = []

for a in range(13):
    for b in range(13):
        cards.append((a, '+', b, a + b))

for a in range(13):
    for b in range(13):
        if a < b:
            continue
        cards.append((a, '-', b, a - b))

# Multiplication and division only
cards = []

for a in range(13):
    for b in range(13):
        cards.append((a, u'\u00d7', b, a * b))

for a in range(1, 13):
    cards.append((a, u'\u00f7', 0, 'inf'))

for b in range(1, 13):
    for ans in range(1, 13):
        a = b * ans
        cards.append((a, u'\u00f7', b, ans))

# If you want a repeatable order, you want to seed the RNG here.
random.shuffle(cards)

card_ids = '0123456789ABCDEFGHIJKLMN'
template_file = "flash-card-template.odt"

sheet = -1
while cards:
    sheet += 1
    print("Generating", sheet)

    page_cards = cards[:24]
    cards = cards[24:]
    front_cards = page_cards[:12]
    back_cards = page_cards[12:]

    doc = desktop.loadComponentFromURL(url_base + template_file, "_blank", 0, ())

    images = doc.getGraphicObjects()
    tables = doc.TextTables
    frames = doc.TextFrames

    replaces = {}

    front_cards.extend(('', '', '', '') * (12 - len(front_cards)))
    back_cards.extend(('', '', '', '') * (12 - len(front_cards)))

    for front_idx, (front, back) in enumerate(zip(front_cards, back_cards)):
        # It would probably be better to simply repeat the code strings in
        # the template twice. Less work here, plus fewer search/replace ops
        # so much faster...
        back_idx = (front_idx + 12) ^ 3
        front_id = card_ids[front_idx]
        back_id = card_ids[back_idx]
        replaces.update({
            "a" + front_id: str(front[0]),
            "b" + front_id: str(front[1]),
            "c" + front_id: str(front[2]),
            "d" + front_id: str(front[3]),
            "e" + front_id: str(back[0]),
            "f" + front_id: str(back[1]),
            "g" + front_id: str(back[2]),
            "a" + back_id: str(back[0]),
            "b" + back_id: str(back[1]),
            "c" + back_id: str(back[2]),
            "d" + back_id: str(back[3]),
            "e" + back_id: str(front[0]),
            "f" + back_id: str(front[1]),
            "g" + back_id: str(front[2]),
        })

    for s, r in replaces.items():
        replace_desc = doc.createReplaceDescriptor()
        replace_desc.SearchString = s
        replace_desc.ReplaceString = r
        xFound = doc.replaceAll(replace_desc)

    print("Saving", sheet)

    file_base = "gen-%03d" % sheet
    doc.storeAsURL(url_base + file_base + ".odt", ())

    # https://bugs.launchpad.net/ubuntu/+source/openoffice.org/+bug/118789 comment 7
    # https://launchpadlibrarian.net/15892085/ooextract-wrap-type-for-exportformfields-property.patch
    filter_data = uno.Any("[]com.sun.star.beans.PropertyValue", (
        createPropertyValue("PageRange", "1-2"),
    ),)
    conv_props = (
        createPropertyValue("FilterName", "writer_pdf_Export"),
        createPropertyValue("FilterData", filter_data),
    )
    doc.storeToURL(url_base + file_base + ".pdf", conv_props)

    doc.close(True)
