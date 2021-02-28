#!/usr/bin/env python
# -*- coding: utf-8 -*-

__author__ = "Max Schmeling"
__copyright__ = "Copyright by Max Schmeling 2019"

__license__ = "GPL" #General Public License
__version__ = "1.0.0"
__maintainer__ = "Max Schmeling"
__status__ = "Development"

"""
NAME: Primary Key Finder
"""

import sys
import os
import time
import itertools
import pandas as pd
from pyxlsb import open_workbook as open_xlsb

EXCELTYPES = (".xls", ".xlsx", ".xlsb", ".csv")
USAGESTRING = "primarykey --hilfe --help --usage | 'path\\to\excel.xlsx' [--worksheet <sheet>] [--range <n> <m>] [--columns <n>] [--precision <p>] [--sort <o>] [--verbose]"


def HelpGerman():
    text = """+---------------------------------------------------------------------+
|        Primärschlüsselfinder (von Max Schmeling im Jahr 2019)       |
+---------------------------------------------------------------------+

 Use --help for the english version of this help message!

BESCHREIBUNG:
 Nimmt eine Tabellendatei und ein Arbeitsblatt und versucht, eine mögliche Spalte 
 oder Kombinationen aus Spalten zu finden, die als Primärschlüssel geeignet
 sind. Wenn keine Spalte die Kriterien eines Primärschlüssels erfüllt (dh. 
 jedes einzelne Element in dieser Spalte darf nur einmal vorhanden sein),
 versucht das Programm einen Primärschlüssel durch Verketten/Kombinieren 
 von maximal <n> Spalten zu finden, wie durch --columns bestimmt wird.
 Wenn es keine Kombinationen von mehreren Spalten gibt, die die Kriterien 
 eines Primärschlüssels erfüllen, dann hat die Tabelle keinen Primärschlüssel.

 Der Parameter --precision kann Pseudo-Primärschlüssel auflisten. Das sind Schlüssel/
 Spalten, in denen <p> % der Elemente nur einmal in dieser Spalte/Spaltenkombination
 erscheinen. Dies kann für Tabellen nützlich sein, in denen kein Primärschlüssel 
 gefunden werden kann, aber wo der Benutzer Einen bestimmten Prozentsatz von nicht
 eindeutigen Tupeln vernachlässigen kann. Wenn das Programm keinen einzigen 
 Primärschlüssel finden konnte werden Vorschläge aufgelistet. Mehr Informationen
 unter --precision.

SYNTAX:
 {0}

PARAMETER:
 --worksheet <sheet> Bestimmt den Namen oder die Nummer des Arbeitsblatts für Excel
                     Dateien. Wenn es nicht angegeben wird, wird standardmäßig das
                     erste Arbeitsblatt der angegebenen Datei verwendet.
 --range <list>      Beschränkt den Scanvorgang auf bestimmte Spalten. <list> muss 
                     eine durch Kommata (,) getrennte Liste ohne Leerzeichen (falls 
                     nicht in Anführungszeichen eingeschlossen) sein. Die Liste kann 
                     Spaltennummern (z. B. 4) und Bereiche (z. B. 5-9) enthalten. 
                     Die Spaltenindizes beginnen bei 1.
 --columns <n>       Die maximale Spaltenmenge, die kombiniert werden kann um einen
                     Primärschlüssel zu genererieren. Der Standardwert ist 3. TIPP: 
                     Je höher <n>, desto wahrscheinlicher ist es, einen Primärschlüssel 
                     zu finden, aber desto mehr Zeit und Rechenleistung wird benötigt.
 --precision <p>     Wenn angegeben, werden am Ende des Scanvorgangs Pseudo-Primärschlüssel 
                     vorgeschlagen. <p> ist ein Prozentwert der den Mindestanteil von Tupeln,
                     die in der jeweiligen Spalte eindeutig sein müssen angibt. Je höher <p>,
                     desto näher ist die Spalte an einem Primärschlüssel. 100% entspricht 
                     einem Primärschlüssel - 0% ist jede Spalte.
 --sort <o>          Sortiert die Ergebnisse abhängig von dem Wert für <o>:
                     1 = Primärschlüssel werden direkt ausgegeben und Pseudo-
                         Primärschlüssel werden am Ende ausgegeben. Pseudo-Primärschlüssel
                         werden nach ihre Präzision sortiert. Nützlich für sofortige
                         und sortierte Ergebnisse. (Standard)
                     2 = Primärschlüssel und Pseudo-Primärschlüssel werden
                         direkt ausgegeben. Pseudo-Primärschlüssel können nicht
                         nach ihrer Präzision sortiert werden. Nützlich für sofortige
                         Ergebnisse zum Beispiel bei großen Dateien.
                     3 = Zeigt einen Fortschrittsbalken an. Primärschlüssel und Pseudo-
                         Primärschlüssel werden am Ende sortiert ausgegeben. Nützlich
                         als Übersicht und für sortierte Ergebnisse.
 --verbose           Druckt jede einzelne Spaltenkombination. Wird ignoriert,
                     wenn --sort 3 aktiv ist.
 --help              Zeigt diese Hilfemeldung in Englisch an.
 --hilfe             Zeigt diese Hilfemeldung an.
 --usage             Zeigt die Befehlssyntax an.
 Jeder Parameter kann mit einem Bindestrich und dem ersten Buchstaben abgekürzt werden.

Unterstützte Dateitypen:
 {1}
""".format(USAGESTRING, ", ".join(EXCELTYPES))
    print(text)
    


def HelpEnglish():
    text = """+---------------------------------------------------------------------+
|            Primary Key Finder (by Max Schmeling in 2019)            |
+---------------------------------------------------------------------+

 Nutze --hilfe für die deutsche Version dieser Hilfsmeldung!

DESCRIPTION:
 Takes a table file and a worksheet and tries to find a column that can be
 used as a primary key. If no column fulfills the criteria of a primary
 key (ie. every single item in that column may only exist once) the 
 program will try to find a primary key by concatenating/combining <n> 
 columns as determined by --columns. If there are NO combinations of 
 multiple columns that fulfill the criteria of a primary key, the 
 given table does NOT have a primary key.
 
 The parameter --precision can list pseudo-primary-keys. Those are keys/
 columns where <p> % of the items only appear once in that [combined] 
 column. This can be useful for tables where no primary key can be 
 found, but where the user can neglect e.g. 0.1% of all items existing
 multiple times. If the program could not find a single primary key
 in a file it will list suggestions automatically. Read --precision 
 for more info.
 
SYNTAX:
 {0}
 
PARAMETERS:
 --worksheet <sheet> Determines the worksheet name or number for excel
                     files. If it is not given it will default to the
                     first worksheet of the given file.
 --range <list>      Limit the scanning process to specific columns.
                     <list> needs to be comma (,) separated list without
                     spaces (if enclose it in quotations). The list
                     can contain column numbers (e.g. 4) and ranges
                     (e.g. 5-9). The column indexes start at 1.
 --columns <n>       The maximum amount columns that can be combined
                     to create a primary key. Defaults to 3. HINT: The 
                     higher <n> the more likely it is to find a primary key, 
                     but the more time and computing power it will take.
 --precision <p>     When given will suggest pseudo-primary-keys at the end
                     of the scanning process. <p> determines the minimum
                     portion of items in a column (or a combination of columns)
                     that need to be unique (ie. only exist once in that
                     column) The higher <p> the closer  the column is to 
                     primary key. 100% is the same as a primary key - 0% 
                     is any column.
 --sort <o>          Sorts the results as determined by <o>:
                     1 = Primarykeys are printed immediately and Pseudo-
                         Primarykeys are printed at the end (Default).
                         Pseudo-Primarykeys will be sorted according to 
                         their precision. Useful for immediate and sorted
                         results.
                     2 = Primarykeys and Pseudo-Primarykeys are printed
                         immediately. Therefore, Pseudo-Primarykeys cannot
                         be sorted according to their precision. Useful 
                         for immediate results.
                     3 = Shows a progressbar. Primarykeys and Pseudo-
                         Primarykeys are printed sorted at the end.
                         Useful as overview and for sorted results.
 --verbose           Prints every single combination of columns. Will be
                     ignored when --sort 3 is enabled.
 --help              Shows this help message.
 --hilfe             Shows help message in German.
 --usage             Shows the command syntax.
 Every paramater can be abbreviated with one dash and its first letter.
 
SUPPORTED FILE TYPES:
 {1}
 """.format(USAGESTRING, ", ".join(EXCELTYPES))
    print(text)
    
    
def ProgressBar(achieved, total, prefix="", suffix=""):
    """
    Creates a visual CLI progress bar by
    visualizing the ratio between the total
    and what has been achieved so far.
    """
    progr_norm = int(achieved/total*20)
    
    block = "█" #█ https://stackoverflow.com/questions/3173320/text-progress-bar-in-the-console
    blocks = progr_norm+1
    
    if blocks > 21:
        blocks = 21
        progr_perc = 100
        progr_norm = 20
        
    space = " "
    post_space = 20-progr_norm
    
    progr_graph = "{}{}".format(blocks*block, post_space*space)
    
    progress = int(achieved/total*100)
    print("{}|{}|{}%| ({} of {} {})".format(prefix, progr_graph, progress, achieved, int(total), suffix), end="\r")
    
    
def ConvertSeconds(time, precision=0):
    """ Converts <time> from seconds to the most appropriate """
    unit = "seconds"
    if time >= 86400: #convert to days
        time = time / 86400
        unit = "days"
    elif time >= 3600: #convert to hours
        time = time / 3600
        unit = "hours"
    elif time >= 60: #convert to minutes
        time = time / 60
        unit = "minutes"
    return [round(time, precision), unit]


def IndexToExcelLetter(index):
    """ Converts column index (starts at 1) to excel's letter syntax """
    letterindex = ""
    while index > 0:
        modulo = (index - 1) % 26
        letterindex = str(chr(65 + modulo)) + letterindex
        index = (index - modulo) // 26
    return letterindex


def PredictCombinations(items, n):
    """
    Calculates the amount of combinations of <items> from
    length 1 to <n> without repition and order.
    """
    def faculty(number):
        faculty = 1
        for i in range(1, number+1):
            faculty = faculty * i
        return faculty
    
    def binomial_coefficient(k, n):
        return faculty(n) / (faculty(n-k)*faculty(k))

    combinations = 0
    for i in range(1, n+1):
        combinations += binomial_coefficient(i, len(items)) 
    return int(combinations)


def ColumnCombinations(items, n):
    """ Wrapper for itertools.combinations() to suit our need """
    for i in range(1, n+1):
        for b in itertools.combinations(items, i):
            yield b
            
            
def ParseColumnIndexes(userstring, _min, _max):
    """ 
    Parses string where user specifies the column indexes to be scanned 
     3-6 = scan columns 3 to 6
     7 = scan column 7
     , = separator
    """
    columnlist = [] # stores the resulting column indexes
    tokenlist = userstring.replace(" ", "").split(",")
    for i, tok in enumerate(tokenlist, 1):
        
        # Process regular numbers. E.g. 5
        if tok.isdigit():
            try:
                col = int(tok)
                if col >= _min and col <= _max:
                    columnlist.append(col)
                else:
                    return ["Listitem {0} in --columns out of range: {1}".format(i, tok)]
            except ValueError:
                # If first element type(str) means error code
                return ["Listitem {0} in --columns is not an int: {1}".format(i, tok)]
        
        # Process range. E.g.: 3-6
        elif "-" in tok and len(tok.split("-")) == 2:
            splittok = tok.split("-")
            try:
                start = int(splittok[0])
                end = int(splittok[1])
                if start < _min:
                    return ["Range start in --columns too low (minimum: {0}): {1}".format(_min, start)]
                if end > _max:
                    return ["Range limiter --columns too high (maximum: {0}): {1}".format(_max, end)]
            except ValueError:
                return ["Listitem {0} in --columns is not a valid range: {1}".format(i, tok)]
            if end < start:
                return []
            elif start == end:
                columnlist.append(start)
            else:
                for i in range(start, end+1):
                    columnlist.append(i)
                    
         # Skip empty listitems
        elif len(tok) == 0 or tok.isspace():
            continue
        
        else:
            return ["Invalid listitem in --columns: {0}. Expected int or range.".format(tok)]
    if len(columnlist) == 0:
        return ["No items after --range parameter. Expected comma separated list.".format(tok)]
    else:
        return sorted(list(set(columnlist))) # create set to get rid of duplicates
    
    
        

def Main(filepath, sheetname, maxcolumns, precision, usercolumns, verbose, sort):
    combinations = 0 # counts combinations
    rowitems = [] # stores rowitems for every column combination. Gets reset for every new column combination.
    colname = "" # holds concatenated row item which are then appended to <rowitems>. Gets reset for every new row.
    primarykeys = 0 # total amount of primary keys found
    pseudokeys = 0 # total amount of pseudo-primary-keys found
    foundkeys = [] # keeps track of all found primary key column index
    primarykeys_results = [] # stores all primary keys so we can print them at the end
    pseudokeys_results = [] # stores all keys that are close to being primary keys so ...
    keystring = "" # holds the (combined) primary keys
    flag = False # used to skip certain keys when the already have been checked
    
    print("Opening", os.path.basename(filepath), "...")
    
    if filepath.endswith(".xlsb"):
        try:
            with open_xlsb(filepath) as wb:
                try:
                    with wb.get_sheet(sheetname) as sheet:
                        for r in sheet.rows(sparse=True):
                            columns = [c.v for c in r]
                            break
                        df = pd.DataFrame([[c.v for c in r] for i, r in enumerate(sheet.rows(sparse=True)) if i > 0], columns=columns)
                except (ValueError, IndexError):
                    print("ERROR: The given worksheet does not exist:", sheetname)
                    print(os.path.basename(filepath), "has the following worksheets:")
                    for i, s in enumerate(wb.sheets, 1):
                        print("{}:\t{}".format(i, s))
                    return (-1, -1) # cancel program code
        except Exception:
            print("ERROR: Could not open file:", filepath)
            return (-1, -1)
    
    if filepath.endswith((".xls", ".xlsx")):
        try:
            if isinstance(sheetname, int):
                df = pd.read_excel(filepath, sheet_name=sheetname-1)
            else:
                df = pd.read_excel(filepath, sheet_name=sheetname)
        except Exception:
            print("ERROR: The given worksheet does not exist:", sheetname)
            print(os.path.basename(filepath), "has the following worksheets:")
            for i, s in enumerate(pd.ExcelFile(filepath).sheet_names, 1): 
                print("{}:\t{}".format(i, s))
            return (-1, -1)
                
    elif filepath.endswith(".csv"):
        try:
            df = pd.read_csv(filepath, sep=None, engine="python") # sep=None and engine="python" lets python guess the separator
        except Exception:
            print("ERROR: Could not open file:", filepath)
            return (-1, -1)
        
        
    # Count columns and columns
    rows, columns = df.shape
    
    # Error handling
    if maxcolumns > columns:
        maxcolumns = columns
    
    
    # Validate columns given by user
    if usercolumns:
        columnlist = ParseColumnIndexes(usercolumns, 1, columns)
        if type(columnlist[0]) == str:
            print("ERROR:", columnlist[0])
            return (-1, -1) # an error occurred during parsing
        columnlist = [c-1 for c in columnlist]
    else:
        columnlist = [i for i in range(0, columns)]
    
    
    # Calculate the amount of column combinations    
    totalcombinations = PredictCombinations(columnlist, maxcolumns)
    
    
    # Print information for user
    if len(columnlist) == columns:
        print("Scanning worksheet '{0}' with {1} columns and {2} rows ...".format(sheetname, len(columnlist), rows))
    else:
        print("Scanning worksheet '{0}' with {1} ({2} selected) columns and {3} rows ...".format(sheetname, columns, len(columnlist), rows))
    if verbose:
        print("Columns to test:", ", ".join(["{}({})".format(str(c+1), IndexToExcelLetter(c+1)) for c in columnlist]))
    print("Testing", totalcombinations, "primary keys ...")
    
    
    # Calculating expected scanning time
    start_time = time.time()
    # The 3 lines below simulate the process of scanning through the columns.
    # They have no actual purpose other than functioning as a dummy to measure
    # the duration of one combination to predict the total scan duraction.
    for testrow in df.itertuples():
        for testcol in [testi for testi in range(maxcolumns)]:
            teststr = testcol
    delta_time = time.time() - start_time
    expected_seconds = delta_time * totalcombinations
    expected = ConvertSeconds(expected_seconds)
    print("Expected scan duration ~ {} {}".format(int(expected[0]), expected[1]))
    print("Starting scan...")
    print()
    
    
    # Create all possible combinations of <maxcolumns> columns and test if they are primary keys
    for keycolumns in ColumnCombinations(columnlist, maxcolumns):
        combinations += 1
        
        # Draw progress bar if the user choosed so
        if sort == 3:
            ProgressBar(combinations, totalcombinations, prefix="", suffix="keys tested")
            if combinations == totalcombinations: print("\n")
            
        # Check if <keycolumns> is a subset of of one of
        # the found PRIMARY-keys, so we can skip it.
        for key in foundkeys:
            for k in key:
                if not k in keycolumns:
                    break
            else:
                _key = key
                flag = True
                break
        
        # if one of the two for-loops above found a match, skip this iteration
        if flag:
            flag = False
            if verbose and sort != 3:
                print("Skipping column(s): ", " + ".join([str(c+1) for c in keycolumns]), "because found", " + ".join([str(c+1) for c in _key]), "already before")
            continue
        
        if verbose and sort != 3:
            print("Testing column(s): ", " + ".join([str(c+1) for c in keycolumns]))
        
        # Create primary keys for each row
        for row in df.itertuples():
            for c in keycolumns:
                keystring += str(row[c+1])
            if len(keystring) > 0:
                rowitems.append(keystring)
            keystring = ""
            
        # Create a new list with all the duplicate rowitems removed.
        setrow = list(set(rowitems))
        
        # If the length of the new list is equal to the total length of rows, a primary key was found
        if len(setrow) == rows:
            primarykeys += 1
            foundkeys.append(keycolumns)
            if sort == 3:
                primarykeys_results.append(["'  +  '".join([str(df.columns[idx]) for idx in keycolumns]),
                                            " + ".join([str(x+1) for x in keycolumns]),
                                            " + ".join([IndexToExcelLetter(idx+1) for idx in keycolumns]),
                                            ])
            else:
                print("Primary Key #{0}:".format(primarykeys))
                print(" Columnname:\t'{0}'".format("'  +  '".join([str(df.columns[idx]) for idx in keycolumns])))
                print(" Columnindex:\t{0}".format(" + ".join([str(x+1) for x in keycolumns])))
                print(" Columnletter:\t{0}".format(" + ".join([IndexToExcelLetter(idx+1) for idx in keycolumns])))
                print()
        elif len(setrow)/rows >= precision: # close to being a primary key. Ie. <precision>*100 (%) of the items are unique
            pseudokeys += 1
            if sort == 2:
                print("Suggestion #{0} (NOT a primary key, but {1}% of items are unique)".format(pseudokeys, round(len(setrow)/rows*100, 8)))
                print(" Columnname:\t'{0}'".format("'  +  '".join([str(df.columns[idx]) for idx in keycolumns])))
                print(" Columnindex:\t{0}".format(" + ".join([str(x+1) for x in keycolumns])))
                print(" Columnletter:\t{0}".format(" + ".join([IndexToExcelLetter(idx+1) for idx in keycolumns])))
                print()
            else:
                pseudokeys_results.append(["'  +  '".join([str(df.columns[idx]) for idx in keycolumns]),
                                         " + ".join([str(x+1) for x in keycolumns]),
                                         " + ".join([IndexToExcelLetter(idx+1) for idx in keycolumns]),
                                         len(setrow)/rows
                                         ])              
        
        # Reset rowitems for next column
        rowitems.clear()
        
    if sort == 3:
        for i, key in enumerate(primarykeys_results, 1):
            print("Primary Key #{0}:".format(i))
            print(" Columnname:\t'{0}'".format(key[0]))
            print(" Columnindex:\t{0}".format(key[1]))
            print(" Columnletter:\t{0}".format(key[2]))
            print()
    
    if primarykeys == 0 and sort == 1 or suggestions:
        pseudokeys_results.sort(key=lambda x: x[3], reverse=True) # sort list in descending order by uniqueness
        for i, key in enumerate(pseudokeys_results, 1):
            if i >= 100: break # Only show the Top-100 pseudo-keys
            print("Suggestion #{0} (NOT a primary key, but {1}% of items are unique)".format(i, round(key[3]*100, 8)))
            print(" Columnname:\t'{0}'".format(key[0]))
            print(" Columnindex:\t{0}".format(key[1]))
            print(" Columnletter:\t{0}".format(key[2]))
            print()
    
    if primarykeys == 0:
        if verbose: print()
        print("No primary key found.\n")
                
    return (primarykeys, pseudokeys) # Return amount of found keys
                

if __name__ == "__main__":
    if len(sys.argv) < 2:          
        print("ERROR: Not enough arguments.")
        print(USAGESTRING)
        sys.exit(1)
        
    if sys.argv[1].lower() in ("-u", "--usage"):
        print(USAGESTRING)
        sys.exit(0)
        
    elif sys.argv[1].lower() in ("--help", "-h", "help", "/?"):
        HelpEnglish()
        sys.exit(0)

    elif sys.argv[1].lower() in ("--hilfe", "hilfe"):
        HelpGerman()
        sys.exit(0)
    
    if not os.path.isfile(sys.argv[1]):
        print("ERROR:", sys.argv[1], "is not a file.")
        sys.exit(1)
        
    if not sys.argv[1].lower().endswith(EXCELTYPES):
        print("ERROR: The given filetype is not supported.")
        print("Supported excel types:", ", ".join(EXCELTYPES))
        sys.exit(1)
    
    
    # Set options (view Help...() for detailed info)
    filepath = sys.argv[1].strip(",").strip('"')
    maxcolumns = 3 # maximum amount of columns when combining columns as a primary key
    usercolumns = None # None means "entire range"
    verbose = False # print detailed info
    sheetname = 1 # default sheetname is the first one
    precision = 0.999 # default precision is 99.9 %
    suggestions = False # make pseudo-primary-key suggestions
    sort = 1 # the order in which the results will be printed
    
    
    # Check Parameters
    for i, arg in enumerate(sys.argv):
        arg = arg.lower()
        
        if arg in ("-w", "--worksheet", "--sheet"):
            try:
                sheetname = int(sys.argv[i+1])
                if sheetname < 1:
                    sheetname = 1
            except Exception:
                sheetname = sys.argv[i+1]
            
        elif arg in ("-c", "--columns"):
            try:
                maxcolumns = int(sys.argv[i+1])
                if maxcolumns < 1:
                    raise Exception
            except Exception:
                print("ERROR: Invalid or no --columns specifier given. Expected integer (>= 1) after", arg, "got {}".format(sys.argv[i+1]))
                sys.exit(0)
            
        elif arg in ("-v", "--verbose"):
            verbose = True
        
        elif arg in ("-p", "--precision"):
            try:
                precision = abs(float(sys.argv[i+1])/100.0)
                suggestions = True
            except Exception:
                print("ERROR: Invalid or no --precision specifier given. Expected percent value after", arg)
                sys.exit(0)
                
        elif arg in ("-r", "--range"):
            try:
                usercolumns = sys.argv[i+1]
            except Exception:
                print("ERROR: Invalid or no --range specifier given. Expected comma separated list after", arg)
                sys.exit(0)
                
        elif arg in ("-s", "--sort"):
            try:
                sort = int(sys.argv[i+1])
            except Exception:
                print("ERROR: Invalid or no --sort specifier given. Expected integer after", arg)
                sys.exit(0)                
                
    # Start process with timer
    start_time = time.time()
    try:
        keys = Main(filepath, sheetname, maxcolumns, precision, usercolumns, verbose, sort)
    except KeyboardInterrupt:
        if sort == 3: print("\n")
        print("Process cancelled through user interaction.")
        print("Thank You for using Primary Key Finder by Max Schmeling!")
        sys.exit(2)
    except Exception as e:
        print("Fuck... an unexpected error occurred:", e)
        sys.exit(1)
    
    if keys[0] < 0:
        sys.exit(1)
    
    end_time = time.time() - start_time #time taken in seconds
    frmt_time = ConvertSeconds(end_time, 2)
    print(frmt_time[0], frmt_time[1], "for", keys[0], "primary key(s) and", keys[1], "suggestion(s)")
    print("Thank You for using Primary Key Finder by Max Schmeling!")
    
    
    