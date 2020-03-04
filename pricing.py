from PySide2 import QtGui, QtCore, QtWidgets
import pandas as pd
import sys, os, time, traceback

class PricingExcelProcessor():
    def process(self, filename, oldfilename):
        try:
            # Validate filename
            if filename.lower().find('.xls') < 0:
                return (False, "ERROR: Only .xls and .xlsx files are supported")
            
            # Open excel file and read relevant sheets
            input_file = open(filename, 'rb')
            fitments = pd.read_excel(input_file, sheet_name="Fitment Group Export File Scrub")
            modified_parts = pd.read_excel(input_file, sheet_name="2. Modified Part Number Table")
            excluded_parts = pd.read_excel(input_file, sheet_name="3. Excluded Part# Table")
            excluded_steps = pd.read_excel(input_file, sheet_name="4. Excluded Step Table")
            valid_speeds = pd.read_excel(input_file, sheet_name="5. Valid Speed Table")
            reformat_sizes = pd.read_excel(input_file, sheet_name="6. Reformat Size Table")
            all_terrain_parts = pd.read_excel(input_file, sheet_name="7. All Terrain Part# Table")
            mud_terrain_parts = pd.read_excel(input_file, sheet_name="8. Mud Terrain Part# Table")
            st_trailer_parts = pd.read_excel(input_file, sheet_name="9. ST Trailer Part# Table")
            input_file.close()

            # Read the old file if exists
            new_parts = None
            try:
                # 1. Create Excel Spreadsheet of new records added since the last upload.
                oldfile = open(oldfilename, 'rb')
                filtered = pd.read_excel(oldfile, sheet_name="Filtered")
                deleted = pd.read_excel(oldfile, sheet_name="Deleted")
                oldfile.close()
                new_parts = fitments[~fitments['Part#'].isin(filtered['Part#']) & ~fitments['Part#'].isin(deleted['Part#'])]
                print('Captured all new parts')

                # Keep and be able to look up spreadsheets from the previous 5 exports.
                newfilename = oldfilename.replace('.xls', '_' + str(time.time()) + '.xls')
                os.rename(oldfilename, newfilename)
            except Exception as e:
                print(e)

            # 2.  Replace Fitment Group Export File values in Model, Tire Size, Section, Aspect, Rim, Load, Speed, Category, and Step Columns with values in the Modified Part Number Table for records with matching Part Numbers.
            fitments = fitments.set_index('Part#')
            modified_parts['Part#'] = modified_parts['Part#'].astype('str')
            modified_parts = modified_parts.drop(columns='Brand').set_index('Part#')
            fitments.update(modified_parts)
            fitments.reset_index(inplace=True)
            del modified_parts
            print("Replaced Fitment Group Export File values")

            #Filter out part numbers
            # 3.  Delete Part#'s where Part# in Excluded Part# Table
            # 4.  Delete Part#'s where Tire Size = "blank"
            # 5.  Delete Part#'s where Section = "blank"
            # 6.  Delete Part#'s where Aspect = "blank"
            # 7.  Delete Part#'s where Rim = "blank"
            # 8.  Delete Part#'s where Load = "blank"
            # 9.  Delete Part#'s where Speed = "blank"
            # 10. Delete Part#'s where Category = "blank"
            # 11. Delete Part#'s where Step = "blank"
            # 12. Delete Part#'s where Step in Excluded Step Table
            # 13. Delete Part#'s where Speed NOT in Valid Speed Table
            ff = fitments[ \
                fitments['Tire Size'].notnull() \
                & fitments['Section'].notnull() \
                & fitments['Aspect'].notnull() \
                & fitments['Rim'].notnull() \
                & fitments['Load'].notnull() \
                & fitments['Speed'].notnull() \
                & fitments['Category'].notnull() \
                & fitments['Step'].notnull() \
                & ~fitments['Part#'].isin(excluded_parts['Part#']) \
                & ~fitments['Step'].isin(excluded_steps['Step']) \
                & fitments['Speed'].isin(valid_speeds['Speed']) \
            ]
            del excluded_parts, excluded_steps, valid_speeds
            print("Filtered out part numbers")

            # 14. Add P to 1st character Tire Size if Category in Reformat Size Table
            # 15. Add LT to 1st character Tire Size if Category in Reformat Size Table
            sizes = pd.merge(ff[['Tire Size', 'Speed', 'Category']], reformat_sizes, \
                how='left').fillna('').set_index([ff.index])
            ff['Tire Size'] = sizes['1st Chacter In Size'].map(str) + sizes['Tire Size']
            del reformat_sizes, sizes
            print("Updated Tire Size - Prefix P & LT")

            # 16. Add ST to 1st character Tire Size if Part# in ST Trailer Part# Table
            st_trailer_parts['Flag'] = 'ST'
            sizes = pd.merge(ff[['Tire Size', 'Part#']], st_trailer_parts[['Part#', 'Flag']], \
                how='left').fillna('').set_index([ff.index])
            ff['Tire Size'] = sizes['Flag'].map(str) + sizes['Tire Size']
            del sizes
            print("Updated Tire Size - Prefix ST")

            # 17. Remove "Z" in Size in Fitment Group Export File if character exists in front of "R"
            ff['Tire Size'] = ff['Tire Size'].str.replace('ZR', 'R')
            print("Updated Tire Size - ZR to R")

            # 18. Replace value in Category in Fitment Group Export File per Logic in Replace Category Table.
            ff.loc[( \
                (ff['Tire Size'].str[0] == 'P') \
                & ~ff['Part#'].isin(all_terrain_parts['Part#']) \
                & ~ff['Part#'].isin(mud_terrain_parts['Part#']) \
                & ~ff['Part#'].isin(st_trailer_parts['Part#']) \
            ), 'Category'] = 'P'
            ff.loc[( \
                (ff['Tire Size'].str[0] == 'P') \
                & ff['Part#'].isin(all_terrain_parts['Part#']) \
                & ~ff['Part#'].isin(mud_terrain_parts['Part#']) \
                & ~ff['Part#'].isin(st_trailer_parts['Part#']) \
            ), 'Category'] = 'PAT'
            ff.loc[( \
                (ff['Tire Size'].str[0] == 'P') \
                & ~ff['Part#'].isin(all_terrain_parts['Part#']) \
                & ff['Part#'].isin(mud_terrain_parts['Part#']) \
                & ~ff['Part#'].isin(st_trailer_parts['Part#']) \
            ), 'Category'] = 'PMT'
            ff.loc[( \
                (ff['Tire Size'].str[:2] == 'LT') \
                & ~ff['Part#'].isin(all_terrain_parts['Part#']) \
                & ~ff['Part#'].isin(mud_terrain_parts['Part#']) \
                & ~ff['Part#'].isin(st_trailer_parts['Part#']) \
            ), 'Category'] = 'LT'
            ff.loc[( \
                (ff['Tire Size'].str[:2] == 'LT') \
                & ff['Part#'].isin(all_terrain_parts['Part#']) \
                & ~ff['Part#'].isin(mud_terrain_parts['Part#']) \
                & ~ff['Part#'].isin(st_trailer_parts['Part#']) \
            ), 'Category'] = 'LTAT'
            ff.loc[( \
                (ff['Tire Size'].str[:2] == 'LT') \
                & ~ff['Part#'].isin(all_terrain_parts['Part#']) \
                & ff['Part#'].isin(mud_terrain_parts['Part#']) \
                & ~ff['Part#'].isin(st_trailer_parts['Part#']) \
            ), 'Category'] = 'LTMT'
            ff.loc[( \
                (ff['Tire Size'].str[:2] == 'LT') \
                & ~ff['Part#'].isin(all_terrain_parts['Part#']) \
                & ~ff['Part#'].isin(mud_terrain_parts['Part#']) \
                & ff['Part#'].isin(st_trailer_parts['Part#']) \
            ), 'Category'] = 'ST'
            del all_terrain_parts, mud_terrain_parts, st_trailer_parts
            print("Updated Category")

            # Capture all dropped fitments
            ex = fitments[~fitments.index.isin(ff.index)]
            print("Captured all dropped fitments")

            # Return the dataframes
            return (True, (ff, ex, new_parts))

        except Exception as e:
            # traceback.print_exc(file=sys.stdout)
            return (False, "ERROR: " + str(e))


class MaddenCoProcessor():
    def process(self, filename):
        try:
            # Validate filename
            if filename.lower().find('.xls') < 0:
                return (False, "ERROR: Only .xls and .xlsx files are supported")
            
            # Open excel file and read relevant sheets
            input_file = open(filename, 'rb')
            madden_co = pd.read_excel(input_file, sheet_name="MddenCo Export Data Scrub")
            input_file.close()

            # 1. Reformat size field If PDCLASS = 01 or 02 or 03 or 04 or 05 or 06 or 07 or 08 then PDSIZE = P185/65R14
            # 2. Reformat size field If PDCLASS = 09 or 10 or 11 then PDSIZE = LT265/70R17
            # 3. Reformat size field If PDCLASS = 13 then PDSIZE = ST225/75R15
            madden_co.loc[madden_co['PDCLASS'].isin(range(1,9)), 'PDSIZE'] = 'P185/65R14'
            madden_co.loc[madden_co['PDCLASS'].isin(range(9,12)), 'PDSIZE'] = 'LT265/70R17'
            madden_co.loc[madden_co['PDCLASS']==13, 'PDSIZE'] = 'ST225/75R15'
            # 4. Replace PDCLASS value with P if PDCLASS = 01 or 02 or 03 or 04 or 05
            # 5. Replace PDCLASS value with PAT if PDCLASS = 06
            # 6. Replace PDCLASS value with PMT if PDCLASS = 07
            # 7. Replace PDCLASS value with LT if PDCLASS = 09
            # 8. Replace PDCLASS value with LTAT if PDCLASS = 10
            # 9. Replace PDCLASS value with LTMT if PDCLASS = 11
            # 10. Replace PDCLASS value with ST if PDCLASS = 13


            # Return the dataframe
            return (True, madden_co)

        except Exception as e:
            return (False, "ERROR: " + str(e))


class FileDnDWidget(QtWidgets.QWidget):
    """
    Subclass the widget and add a button to load files. 
    
    Alternatively set up dragging and dropping of files onto the widget
    """
    def __init__(self, main, title):
        super(FileDnDWidget, self).__init__()

        # Accept input arguments: main window and widget title
        self.main = main
        self.title = title

        # Button that allows loading of files
        self.load_button = QtWidgets.QPushButton("Open {} Excel File".format(title))
        self.load_button.setFixedWidth(360)
        self.load_button.clicked.connect(self.load_file_btn)

        # Drag & Drop region
        self.dnd_label = QtWidgets.QLabel(self)
        self.dnd_label.setFixedWidth(360)
        self.dnd_label.setAlignment(QtCore.Qt.AlignCenter)
        self.dnd_label.setText("OR Drop Here")

        # Set the widget layout
        self.layout = QtWidgets.QVBoxLayout()
        self.layout.addWidget(self.load_button)
        self.layout.addWidget(self.dnd_label)
        self.setLayout(self.layout)

        # Allow files to be dragged and dropped
        self.setAcceptDrops(True)

    def load_file_btn(self):
        """
        Open a File dialog when the button is pressed
        """        
        #Get the file location
        fileDialog = QtWidgets.QFileDialog()
        fileDialog.setFileMode(fileDialog.ExistingFile)
        fname, _ = fileDialog.getOpenFileName(self, 'Open file', filter='Excel Files (*.xls *.xlsx)')
        # Load the file from the location
        if fname:
            self.load_file(fname)

    def load_file(self, fname):
        """
        Set the filename in main window for processing
        """
        self.main.setFileName(self.title, fname)

    def isAcceptable(self, e):
        return e.mimeData().hasUrls()

    def acceptFile(self, e):
        if self.isAcceptable(e):
            e.accept()
        else:
            e.ignore()

    # The following three methods set up dragging and dropping for the app
    def dragEnterEvent(self, e):
        self.acceptFile(e)

    def dragMoveEvent(self, e):
        self.acceptFile(e)

    def dropEvent(self, e):
        """
        Drop files directly onto the widget
        Last file's path is stored in fname
        """
        if self.isAcceptable(e):
            e.setDropAction(QtCore.Qt.CopyAction)
            e.accept()
            fname = ''
            for url in e.mimeData().urls():
                fname = str(url.toLocalFile())
            self.load_file(fname)
        else:
            e.ignore()


class MainWindowWidget(QtWidgets.QWidget):
    """
    Subclass the widget and add a button to load files. 
    
    Alternatively set up dragging and dropping of files onto the widget
    """

    def __init__(self):
        super(MainWindowWidget, self).__init__()
        self.pricing_file = False
        self.maddenco_file = False
        
        # Create an Excel File Process instance
        self.pricing_xl_processor = PricingExcelProcessor()
        self.madden_co_processor = MaddenCoProcessor()

        # Drag & Drop Regions for loading files
        self.pricing_file_region = FileDnDWidget(self, 'Pricing')
        self.maddenco_file_region = FileDnDWidget(self, 'MaddenCo')

        # Create Process & Reload buttons
        self.process_button = QtWidgets.QPushButton("Process Opened Files")
        self.process_button.clicked.connect(self.process_files)
        self.reload_button = QtWidgets.QPushButton("Process More Files")
        self.reload_button.clicked.connect(self.reload_gui)

        # Create status message box layout
        self.status_label = QtWidgets.QLabel(self)
        self.status_label.setFixedWidth(720)
        self.status_label.setAlignment(QtCore.Qt.AlignCenter)

        # Set the window strating layout
        self.regionlayout = QtWidgets.QHBoxLayout()
        self.regionlayout.addWidget(self.pricing_file_region)
        self.regionlayout.addWidget(self.maddenco_file_region)
        self.layout = QtWidgets.QVBoxLayout()
        self.layout.addLayout(self.regionlayout)
        self.layout.addWidget(self.process_button)
        self.layout.addWidget(self.reload_button)
        self.layout.addWidget(self.status_label)
        self.setLayout(self.layout)

        self.reload_gui()
        self.setWindowTitle('Pricing Application')
        self.show()

    def reload_gui(self):
        # self.setAcceptDrops(True)
        self.reload_button.hide()
        self.process_button.show()
        self.pricing_file_region.show()
        self.maddenco_file_region.show()
        self.pricing_file_region.setAcceptDrops(True)
        self.maddenco_file_region.setAcceptDrops(True)
        self.adjustSize()
        self.pricing_file = False
        self.maddenco_file = False
        self.status_label.setText("Drag and drop .xls or .xlsx files on to application window, or open using Open button...")

    def setFileName(self, title, fname):
        ftype = title.lower()
        if ftype == 'pricing':
            self.pricing_file = fname
            self.status_label.setText("Opened Pricing File...")
        else:
            self.maddenco_file = fname
            self.status_label.setText("Opened MaddenCo File...")

    def process_files(self):
        # Validate pricing file load
        if not self.pricing_file:
            self.status_label.setText("Please Open Pricing File")
            return

        # Validate maddenco file load
        if not self.maddenco_file:
            self.status_label.setText("Please Open MaddenCo File")
            return

        # Disable drag & drop
        self.pricing_file_region.setAcceptDrops(False)
        self.maddenco_file_region.setAcceptDrops(False)
        self.pricing_file_region.hide()
        self.maddenco_file_region.hide()

        # Update processing status
        self.status_label.setText("Processing Excel Files...")
        self.process_button.hide()
        self.adjustSize()
        QtWidgets.qApp.processEvents()
        
        """
        Process the files
        """
        # Get output filename
        outname = self.pricing_file.replace('.xls', '_scrubbed.xls')

        # Process pricing file
        success, dataframes = self.pricing_xl_processor.process(self.pricing_file, outname)
        if not success:
            self.status_label.setText(dataframes)
            self.reload_button.show()
            return
        self.status_label.setText("Processed Pricing File")
        QtWidgets.qApp.processEvents()
        fitment, discarded, new_parts = dataframes

        # Process maddenco file
        success, madden_co = self.madden_co_processor.process(self.maddenco_file)
        if not success:
            self.status_label.setText(madden_co)
            self.reload_button.show()
            return
        self.status_label.setText("Processed MaddenCo File")
        QtWidgets.qApp.processEvents()

        # Write to output file
        with pd.ExcelWriter(outname) as writer:
            fitment.to_excel(writer, sheet_name="Filtered", index=False)
            discarded.to_excel(writer, sheet_name="Deleted", index=False)
            if new_parts is not None:
                new_parts.to_excel(writer, sheet_name="New Part Numbers", index=False)
            madden_co.to_excel(writer, sheet_name="MaddenCo", index=False)

        # Update the status and GUI
        self.status_label.setText("Scrubbed data has been written to " + outname)
        self.reload_button.show()
        # self.adjustSize()

# Run if called directly
if __name__ == '__main__':
    pd.options.mode.chained_assignment = None
    # Print usage info in any case
    print("Drag and drop .xls or .xlsx files on to application window, or open using Open button...")
    # Initialise the application
    app = QtWidgets.QApplication(sys.argv)
    # Call the widget
    ex = MainWindowWidget()
    sys.exit(app.exec_())