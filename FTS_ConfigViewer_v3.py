# FTS Configuration Viewer
# ottO Bédard, MSc.
# Environment and Climate Change Canada
# otto.bedard@canada.ca
# Developed for Water Survey of Canada for ease of reading configuration files from FTS H1/H2 data loggers

# Version History
# v3.0 2026-01-13
# Complete overhaul to PyQt5 with modern UI, file pickers, and modular functions.
# added pdf generation
# no longer requires internet access for XSLT as it uses local copy

# v2.5 2020-12-04
# changed local .xsl to github repository for FireFox fix and long term adaptability.

# v2.3/4 added none-FTP version checking
# 2019-12-04

# v 2.2 # added Menu List and Version Checking capability

import os
import sys
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget,
    QVBoxLayout, QLabel, QPushButton, QMessageBox,
    QFileDialog, QAction
)
from PyQt5.QtGui import QPixmap, QFont, QIcon, QDesktopServices
from PyQt5.QtCore import Qt, QUrl

# V3 stuff
import pdfkit
from lxml import etree

# Import from the pdf generation functions
from pdfGenerationFunctions import xml_to_html, html_to_pdf, xml_to_pdf


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # Set up the main window properties
        self.setWindowTitle("FTS Config Viewer")
        self.setGeometry(100, 100, 300, 400)  # Window position and size (x, y, width, height)
        # Set the window icon
        self.setWindowIcon(QIcon("icon.ico"))  # Replace "icon.ico" with your icon file name
        self.setStyleSheet("""
            QMainWindow {
                background-color: white;
            }
            QLabel {
                color: darkblue;
                font-size: 20px;
            }
            QPushButton {
                background-color: lightgray;
                border: 1px solid darkgray;
                border-radius: 5px;
                padding: 5px;
            }
            QPushButton:hover {
                background-color: gray;
                color: white;
            }"""
        )

        # Call the method to set up the UI layout
        self.setup_ui()


    # Method to set up the UI layout
    def setup_ui(self):
        # Central widget holds all UI components
        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)

        # Create a vertical layout
        layout = QVBoxLayout()

        # Add a title label
        title_label = QLabel("FTS Config Viewer v3", self)
        title_label.setFont(QFont("Arial", 28, QFont.Bold))  # Customize font
        title_label.setAlignment(Qt.AlignCenter)  # Center-align the label
        layout.addWidget(title_label)  # Add title to the layout
        layout.addStretch(1)  # Add stretch to separate title from buttons

        # Add a logo (optional)
        logo_label = QLabel()
        pixmap = QPixmap("wsc.gif")  # Replace "logo.png" with your image file
        if not pixmap.isNull():
            pixmap = pixmap.scaled(125, 125, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            logo_label.setPixmap(pixmap)
            logo_label.setAlignment(Qt.AlignCenter)
            layout.addWidget(logo_label)  # Add logo to layout
        else:
            print("Logo image not found. Please place 'logo.png' in the working directory.")
        layout.addStretch(1)

        # Add buttons
        self.create_button(layout, "Load Logger XML", self.load_logger_xml)
        self.create_button(layout, "Load End Visit Report", self.load_end_visit_report)
        self.create_button(layout, "Generate PDF", self.generate_pdf)
        self.create_button(layout, "Reset", self.reset)
        self.create_button(layout, "Exit", self.close_application)

        # Set the layout
        central_widget.setLayout(layout)

        # Create a menu bar
        menubar = self.menuBar()

        # Create "Help" menu
        help_menu = menubar.addMenu("Help")

        # Add "Online Help" action
        help_action = QAction("Confluence", self)
        help_action.triggered.connect(self.open_help_url)
        help_menu.addAction(help_action)        

    # Helper method to create buttons
    def create_button(self, layout, text, function):
        """
        Helper method to create buttons and add them to the layout.
        """
        button = QPushButton(text)
        button.setFont(QFont("Arial", 12))  # Customize button font/style
        button.clicked.connect(function)  # Connect button to function
        layout.addWidget(button)  # Add button to the layout

    def open_help_url(self):
        """
        Open a web browser to display the program's help wiki.
        """
        url = QUrl("https://watersurveyofcanada.atlassian.net/wiki/spaces/WSCan/pages/616366098/FTS+Config+Viewer+Software")
        if not QDesktopServices.openUrl(url):
            QMessageBox.warning(self, "Error", f"Could not open {url.toString()}.")

    # load XML function
    def load_logger_xml(self):
        """
        Prompts the user to select an .xml file using a file picker dialog.
        """
        # Open a file dialog to select an XML file
        file_path, _ = QFileDialog.getOpenFileName(
            self,                      # Parent widget (the main window)
            "Select Logger XML File",  # Dialog title
            "",                        # Starting directory (empty string defaults to current directory)
            "XML Files (*.xml);;All Files (*)"  # File filter (to show only .xml files by default)
        )

        if file_path:  # If a file was selected
            # QMessageBox.information(self, "File Selected", f"You selected: {file_path}")
            # Do something with the .xml file path, e.g., store it or process it
            self.logger_xml_path = file_path  # Example: Save the file path to a class variable
        else:
            QMessageBox.warning(self, "No File Selected", "Please select a valid XML file.")

    # Load EVR function
    def load_end_visit_report(self):
        """
        Prompts the user to select an End Visit Report using a file picker dialog.
        """
        if not getattr(self, 'logger_xml_path', None):  # Check if logger_xml_path is empty or not set
            QMessageBox.warning(self, "Action Blocked", "Please load a Logger XML file first!")
            return

        # Open a file dialog to select an XML file
        file_path, _ = QFileDialog.getOpenFileName(
            self,                      # Parent widget (the main window)
            "Select End Visit Report",  # Dialog title
            "",                        # Starting directory (empty string defaults to current directory)
            "EVR Files (*.txt);;All Files (*)"  # File filter (to show only .txt files by default)
        )

        if file_path:  # If a file was selected
            # QMessageBox.information(self, "File Selected", f"You selected: {file_path}")
            # Do something with the .xml file path, e.g., store it or process it
            self.logger_evr_path = file_path  # Example: Save the file path to a class variable
        else:
            QMessageBox.warning(self, "No File Selected", "Please select a valid End Visit Report.")

    # Generate PDF function
    def generate_pdf(self):
        """
        Generate the PDF from the loaded XML and EVR files.
        
        :param self: Description
        """
        if not getattr(self, 'logger_xml_path', None):  # Check if logger_xml_path is empty or not set
            QMessageBox.warning(self, "Action Blocked", "Please load a Logger XML file first!")
            return
        
        # Check for the C:\Temp directory, if not make it
        # We are storing the files the user sees in this directory
        if not os.path.isdir('C:\\temp\\FTSViewer\\'):
            # Let's make a dir
            print('Creating C:\\temp\\FTSViewer\\')
            os.makedirs('C:\\temp\\FTSViewer\\')

        # Step #1: Determine if we can an EVR file
        if not getattr(self, 'logger_evr_path', None):  # Check if logger_evr_path is empty or not set
            print("No EVR file loaded, proceeding with XML only.")
            evr_path = ""
        else:
            evr_path = self.logger_evr_path
            print(f"Using EVR file: {evr_path}")
                    # Call the EVR Function
        
            # Read in the End Visit Report EVR
            print('Attempting to Open: ' + evr_path)
        
            # list of information that we want out of the file
            searchitems = ['Logger Model:','Logger Version:','Serial Number:','OS Version:','Software Version:','Serial#:','SW Ver:','Device Type:','Standard:','Antenna Bearing:','Antenna Inclination:']
        
            # make a output file for when we find the test we want
            evr_output=[]

            # Point the program to the file, and open it for reading
            with open(evr_path, "r") as f:
                #Read in the file
                searchlines = f.readlines()
        
            # a simple counter
            count = 0

            # Cycle through the lines of the file
            for i, line in enumerate(searchlines):
            
                # Look for any of the words in the list
                for word in searchitems:
        
                    if word in line: # We found one of the search items
                
                        # Build up the output, without the \n, for the final output
                        # we are going to develop some xml here
                        current_output = line.split(":",1)
                        leftout = current_output[0]
                    
                        # We have to do some extra in case the word has #
                        # if not, just do it regularly
                        leftout = leftout.replace("Serial#","Serial")
                    
                        # It is possible to have two instances of "Device Type" and xml doesn't like that
                        if leftout.find("Device Type") != -1:
                            # found one, up the count
                            count += 1
                            #print (count)

                            if count==2:
                                # We have a second instance of Device Type
                                # add a 2 to the end of the device type
                                leftout=leftout + '2'

                        # create the attribute information
                        rightout = "\"" + current_output[1].strip() + "\" "

                        # Add the tag to the file
                        evr_output.append(leftout.replace(" ","") + "=" + rightout)

            # Close the file
            f.close()
            print("EVR Data Extracted:")
        
        # Step #2: Read in the original XML file
        xml_file = open(self.logger_xml_path, 'r')
        # Skip the first two lines - we add them later!
        xml_file.readline()
        xml_file.readline()
        xml_content = xml_file.read() # Read the rest of the file
        xml_file.close() # Close the file


        newfilename = 'C:\\temp\\FTSViewer\\' + os.path.basename(self.logger_xml_path)
        newfilename = newfilename[0:len(newfilename)-4]
        newfilename = newfilename + "_FTSConfigViewer.xml"
        print('Creating newfile: ' + newfilename)
        newfile = open(newfilename,'w')
        
        # Step 2b: add the XML Declaration and XSLT reference
        newfile.write('<?xml version="1.0" encoding="utf-8"?><XMLRoot>')
        
        # Step 3: add the EVR to the newfile
            # If there is NO EVR LOADED, SKIP THIS!!!

        if os.path.isfile(evr_path):
            # Append the File with the NEW data
            newfile.write('<VisitReport ')
            newfile.write(''.join(evr_output))
            newfile.write('></VisitReport>')

        # Step 4: Append the original XML content
        newfile.write(xml_content)  # Append the original XML content
        newfile.close() # Close the new file

        
        # Step 5: Now generate the PDF
        try:    
            pdf_path = os.path.basename(self.logger_xml_path)[0:len(os.path.basename(self.logger_xml_path))-4] + '_FTSConfigViewer.pdf'
            xml_to_pdf(newfilename, "FTSConfigViewer.xsl", 'C:\\temp\\FTSViewer\\' + pdf_path)
            os.startfile('C:\\temp\\FTSViewer\\' + pdf_path) # Open the generated PDF file

        except Exception as e:
            print("Error generating PDF:", e)

    # Reset function    
    def reset(self):
        """
        Resets the file paths for Logger XML and End Visit Report.
        """
        # Clear the paths
        self.logger_xml_path = ""
        self.end_visit_report_path = ""
        # Provide feedback to the user
        QMessageBox.information(self, "Reset Successful", "Reset successful. You can load new files now.")

    # Close application function
    def close_application(self):
        # Show confirmation dialog before exiting
        reply = QMessageBox.question(self, "Exit Application", "Are you sure you want to exit?",
                                      QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            QApplication.quit()  # Quit the application

# Boilerplate code to run the PyQt application
if __name__ == "__main__":
    app = QApplication(sys.argv)

    window = MainWindow()
    window.show()

    sys.exit(app.exec_())
