from PyQt5.QtWidgets import QMainWindow, QAbstractItemView, QTableWidgetItem, QMessageBox, QComboBox, QApplication, \
    QFileDialog, QDateEdit
from PyQt5 import uic
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import QDate
from itinerary import Itinerary
from report import results
from data import businesses_dict, parish, grade, area, period_start, period_end, datetime
import glob
import json
import logo_rc
import os


class MainWindow(QMainWindow):
    def __init__(self, itinerary: Itinerary):
        super(MainWindow, self).__init__()
        uic.loadUi("ui/nis_inspector.ui", self)

        # Window Title
        self.setWindowTitle("NIS Inspector")

        # Window Icon
        self.setWindowIcon(QIcon("ui/NIS_ICON.ico"))

        # Fixed size
        self.setFixedSize(950, 700)

        # Itinerary class
        self.itinerary = itinerary

        # Initial inputs for Itinerary inputs
        self.officer_grade_combo.addItems(grade)
        self.parish_office_combo.addItems(parish)
        self.area_combo.addItems(area)
        self.period_start.setDate(period_start)
        self.period_end.setDate(period_end)

        # Initialize week ended date on report
        self.week_ended_date.setDate(period_end)

        # Initial inputs for objectives
        visits = list(self.itinerary.report.results_total.keys())
        self.visit_combo.addItems(visits)
        self.business_name_combo.addItems(businesses_dict["NAME_OF_EMPLOYER"])
        self.date_entry.setDate(QDate.currentDate())

        # Submit objective button
        self.submit_obj_btn.clicked.connect(self.get_objective_info)

        # Open document button
        self.open_document_btn.clicked.connect(self.open_document)

        # Save itinerary document
        self.print_itinerary_btn.clicked.connect(self.print_itinerary)

        # Save report document
        self.print_report_btn.clicked.connect(self.print_report)

        # Remove row button
        self.remove_row_btn.clicked.connect(self.remove_row)

        # Reset all the tables
        self.reset_table_btn.clicked.connect(self.reset_tables)

        # Submit results button
        self.update_results_btn.clicked.connect(self.submit_result)

        # Load file button
        self.load_itinerary_btn.clicked.connect(self.load_file)

        # Table edits
        self.itinerary_table.horizontalHeader().setVisible(True)
        self.report_table.horizontalHeader().setVisible(True)

        # Detect table edits
        self.itinerary_table.cellChanged.connect(self.cell_update)
        self.report_table.cellChanged.connect(self.cell_update)

        # Files section
        self.files_onhand_entry.setEnabled(False)
        self.total_entry.setEnabled(False)

        # Load from autosave
        self.load_autosave()

        self.show()

    def load_autosave(self):
        default_dir = os.path.expanduser("~/Desktop/NIS Saves/auto_saves")
        last_save_list = glob.glob(f"{default_dir}/*.json")

        if not last_save_list:
            return
        try:
            with open(last_save_list[-1], "r") as file:
                data = json.load(file)
            self.load_data(data)
        except Exception:
            pass

    @staticmethod
    def show_popup(title, text, alert_type):
        msg = QMessageBox()
        msg.setWindowTitle(title)
        msg.setText(text)
        msg.setWindowIcon(QIcon("ui/NIS_ICON.ico"))
        if alert_type == "info":
            msg.setIcon(QMessageBox.Information)
        elif alert_type == "error":
            msg.setIcon(QMessageBox.Critical)
        msg.exec()

    # Open the folder with save files and launch it from there.
    def open_document(self):
        self.create_files_folder()  # Initialize folders
        default_dir = os.path.expanduser("~/Desktop/NIS Saves/")
        # Open file dialog
        filename = QFileDialog.getOpenFileName(self, "Open File", default_dir, "Word files (*.docx)")
        if filename[0] == "" or filename is None:
            return

        # Launch the document from the dialog box
        os.startfile(filename[0])

    def load_file(self):
        self.create_files_folder()  # Initialize folders
        default_dir = os.path.expanduser("~/Desktop/NIS Saves/state_saves")
        filename = QFileDialog.getOpenFileName(self, "Open File", default_dir, "JSON files (*.json)")
        if filename[0] == "" or filename is None:
            return

        # print(filename[0])
        with open(filename[0], "r") as file:
            data = json.load(file)

        # File error exception

        self.load_data(data)
        # print("file_loaded:", filename[0])

    def load_data(self, data):
        itinerary = self.itinerary
        report = self.itinerary.report

        # Get all info from dictionary
        itinerary.name = data["name"]
        itinerary.grade = data["grade"]
        itinerary.period_start = data["period_start"]
        itinerary.period_end = data["period_end"]
        itinerary.area = data["area"]
        itinerary.parish_office = data["parish_office"]
        report.week_ended = data["week_ended"]
        report.files_carried_fwd = data["files_carried_fwd"]
        report.files_received = data["files_received"]
        report.total = data["total"]
        report.files_cleared = data["files_cleared"]
        report.files_on_hand = data["files_on_hand"]
        report.total_effective = data["total_effective"]
        report.total_ineffective = data["total_ineffective"]
        report.results_total = data["results_total"]
        itinerary.objectives = data["objectives"]

        # Update itinerary and report tables
        self.update_itinerary_table()

        # Set files info
        self.prev_work_entry.setText(data["files_carried_fwd"])
        self.received_entry.setText(data["files_received"])
        self.files_cleared_entry.setText(data["files_cleared"])

        # Update Period
        period_start_2 = datetime.datetime.strptime(data["period_start"], '%d %b %Y')
        period_end_2 = datetime.datetime.strptime(data["period_end"], '%d %b %Y')
        self.period_start.setDate(period_start_2)
        self.period_end.setDate(period_end_2)
        self.week_ended_date.setDate(period_end_2)

        if self.prev_work_entry.text() == "" or self.received_entry.text() == "" or self.files_cleared_entry.text() == "":
            self.prev_work_entry.setText("0")
            self.received_entry.setText("0")
            self.files_cleared_entry.setText("0")
        # Set values for the results
        self.set_combo_value(self.report_table, 5, 'result')
        # Update results
        self.submit_result()

    def cell_update(self, row, col):
        key = ["date", "business_name", "business_address", "visit", "ref_no"]
        report = self.itinerary.report.objectives

        if col <= 3:
            item = self.itinerary_table.item(row, col)
            report[row][key[col]] = item.text()
            # Auto update business info if the business in the database
            business_name = report[row]["business_name"]

            # Check if business_name exists and update the corresponding reference number and address
            try:
                ref_no = self.get_business_info(report[row]["business_name"], "ref_no")
                address = self.get_business_info(report[row]["business_name"], "address")
            except ValueError:
                report[row]["ref_no"] = "-"
            else:
                report[row]["ref_no"] = ref_no
                report[row]["business_address"] = address

        elif col == 4:
            # update the reference number
            item = self.report_table.item(row, col)
            report[row][key[col]] = item.text()

        # Update all the tables
        self.update_itinerary_table()
        self.autosave()

    def reset_tables(self, data):
        # Initialize objectives
        self.itinerary.objectives = []

        # Update itinerary and report tables
        self.update_itinerary_table()

    def get_heading(self):
        # Itinerary
        self.itinerary.name = self.officer_name_input.text()
        self.itinerary.grade = self.officer_grade_combo.currentText()
        self.itinerary.period_start = self.period_start.text()
        self.itinerary.period_end = self.period_end.text()
        self.itinerary.area = self.area_combo.currentText()
        self.itinerary.parish_office = self.parish_office_combo.currentText()

        # Report
        self.itinerary.report.officer = self.itinerary.name
        self.itinerary.report.area = self.itinerary.area

    # Get the corresponding business address or reference number for a business name
    @staticmethod
    def get_business_info(business_name, option):
        index = businesses_dict["NAME_OF_EMPLOYER"].index(business_name)
        if option == "address":
            return businesses_dict["ADDRESS"][index]
        if option == 'ref_no':
            return businesses_dict["REF_NO"][index]
        return

    def get_objective_info(self):
        obj = {}
        # Get business info from database
        business_name = self.business_name_combo.currentText()
        business_address = self.get_business_info(business_name, "address")
        ref_no = self.get_business_info(business_name, "ref_no")

        # Update the object dictionary with all the info
        obj["date"] = self.date_entry.text()
        obj["business_name"] = business_name
        obj["business_address"] = business_address
        obj["ref_no"] = ref_no
        obj["visit"] = self.visit_combo.currentText()
        obj['result'] = "effective"

        # Update the itinerary object with the objectives
        self.itinerary.objectives.append(obj)
        self.update_itinerary_table()

    def update_itinerary_table(self):
        self.itinerary_table.blockSignals(True)
        # Pre-set number of rows
        number_of_objectives = len(self.itinerary.objectives)
        self.itinerary_table.setRowCount(number_of_objectives)
        row = 0
        # Update table row by row
        for obj in self.itinerary.objectives:
            self.itinerary_table.setItem(row, 0, QTableWidgetItem(obj["date"]))
            self.itinerary_table.setItem(row, 1, QTableWidgetItem(obj["business_name"]))
            self.itinerary_table.setItem(row, 2, QTableWidgetItem(obj["business_address"]))
            self.itinerary_table.setItem(row, 3, QTableWidgetItem(obj["visit"]))
            row += 1
        # Update report table
        self.update_report_table()
        # Update the report object with the itinerary information
        self.itinerary.update_report()
        self.itinerary_table.blockSignals(False)

    def update_itinerary_table_2(self):
        self.itinerary_table.blockSignals(True)
        # Pre-set number of rows
        number_of_objectives = len(self.itinerary.objectives)
        self.itinerary_table.setRowCount(number_of_objectives)
        visits = list(self.itinerary.report.results_total.keys())
        row = 0
        # Update table row by row
        for obj in self.itinerary.objectives:
            # Date of objective
            date_edit = QDateEdit()
            date_edit.setDisplayFormat("dd MMM yyyy")
            dt = datetime.datetime.strptime(obj["date"], '%d %b %Y')
            date_edit.setDate(dt)

            # Type of visit
            visit_combo = QComboBox()
            visit_combo.addItems(visits)
            visit_combo.setCurrentText(obj["visit"])

            self.itinerary_table.setCellWidget(row, 0, date_edit)
            self.itinerary_table.setItem(row, 1, QTableWidgetItem(obj["business_name"]))
            self.itinerary_table.setItem(row, 2, QTableWidgetItem(obj["business_address"]))
            self.itinerary_table.setCellWidget(row, 3, visit_combo)
            row += 1
        # Update report table
        self.update_report_table()
        # Update the report object with the itinerary information
        self.itinerary.update_report()
        self.itinerary_table.blockSignals(False)

    def remove_row(self):
        # Remove row from both tables and information from objects
        if self.itinerary_table.rowCount() > 0:
            current_row = self.itinerary_table.currentRow()
            self.itinerary_table.removeRow(current_row)
            self.report_table.removeRow(current_row)
            del self.itinerary.objectives[current_row]
            self.itinerary.update_report()
            self.autosave()

    def update_report_table(self):
        self.report_table.blockSignals(True)
        # Pre-set number of rows
        number_of_objectives = len(self.itinerary.objectives)
        self.report_table.setRowCount(number_of_objectives)
        row = 0
        # Update table row by row
        for obj in self.itinerary.objectives:
            results_combo = QComboBox()
            results_combo.addItems(results)
            self.report_table.setItem(row, 0, QTableWidgetItem(obj["date"]))
            self.report_table.setItem(row, 1, QTableWidgetItem(obj["business_name"]))
            self.report_table.setItem(row, 2, QTableWidgetItem(obj["business_address"]))
            self.report_table.setItem(row, 3, QTableWidgetItem(obj["visit"]))
            self.report_table.setItem(row, 4, QTableWidgetItem(obj["ref_no"]))
            self.report_table.setCellWidget(row, 5, results_combo)
            row += 1
        self.get_heading()
        self.autosave()
        self.report_table.blockSignals(False)

    def set_combo_value(self, table, col, key):
        report = self.itinerary.report.objectives
        row = 0
        for obj in report:
            index = table.model().index(row, col)
            widget = table.indexWidget(index)
            if key == 'result':
                if isinstance(widget, QComboBox):
                    widget.setCurrentText(obj[key].capitalize())
            else:
                if isinstance(widget, QComboBox):
                    widget.setCurrentText(obj[key])
            row += 1

    def submit_result(self):
        # Check the if all the cells are filled
        if self.prev_work_entry.text() == "" or self.received_entry.text() == "" or self.files_cleared_entry.text() == "":
            self.show_popup("Field is required", "You left a field empty and a value must be entered.", "info")
            return False

        row_count = self.itinerary_table.rowCount()
        # Update the results from the report table to the report object
        for row in range(row_count):
            result = self.get_results_combo_value(row, 5)
            self.itinerary.report.objectives[row]['result'] = result
        # Calculate the sum of the results in the report object
        self.itinerary.report.objective_results()
        # Update the summary table in the report
        self.update_summary()
        # Get the files information from the entries and update the object
        self.get_files_info()
        # Get week ended date
        self.itinerary.report.week_ended = self.week_ended_date.text()
        self.get_heading()
        self.autosave()
        return True

    def get_results_combo_value(self, row, col):
        # Get the result combo value in the report table
        index = self.report_table.model().index(row, col)
        widget = self.report_table.indexWidget(index)
        if isinstance(widget, QComboBox):
            current_text = widget.currentText()
            return current_text.lower()

    def update_summary(self):
        visit = self.itinerary.report
        col = 0

        # For each row in the results table, update the info on effective, ineffective and total
        for key in self.itinerary.report.results_total:
            self.summary_table.setItem(0, col, QTableWidgetItem(str(visit.effective[key])))
            self.summary_table.setItem(1, col, QTableWidgetItem(str(visit.ineffective[key])))
            self.summary_table.setItem(2, col, QTableWidgetItem(str(visit.results_total[key])))
            col += 1
        self.summary_table.setItem(0, 6, QTableWidgetItem(str(visit.effective['TOTAL'])))
        self.summary_table.setItem(1, 6, QTableWidgetItem(str(visit.ineffective['TOTAL'])))

    def get_files_info(self):
        # Get file info from the entries
        files_carried_fwd = self.prev_work_entry.text()
        files_received = self.received_entry.text()
        files_cleared = self.files_cleared_entry.text()

        # Calculate the information
        total = int(files_carried_fwd) + int(files_received)
        files_on_hand = total - int(files_cleared)

        # Update the fields in the UI
        self.total_entry.setText(str(total))
        self.files_onhand_entry.setText(str(files_on_hand))

        # Update report
        self.itinerary.report.files_carried_fwd = files_carried_fwd
        self.itinerary.report.files_received = files_received
        self.itinerary.report.total = total
        self.itinerary.report.files_cleared = files_cleared
        self.itinerary.report.files_on_hand = files_on_hand

    @staticmethod
    def reformat_date(date: object) -> object:
        dt = datetime.datetime.strptime(date.text(), '%d %b %Y')
        period = datetime.date.strftime(dt, "%d-%b-%Y")
        return period

    def print_itinerary(self):
        # Get all the heading info from UI and update report and itinerary objects
        self.get_heading()
        self.save_to_document("itinerary", save_type="itinerary")

    def print_report(self):
        # Check if field are empty in reports
        if self.submit_result() is False:
            return
        # Get all the heading info from UI and update report and itinerary objects
        self.get_heading()
        self.save_to_document("report", save_type="report")

    def save_to_document(self, name, save_type=None):
        filename = f"{name}-{self.reformat_date(self.period_start)}.docx"
        default_dir = os.path.expanduser("~/Desktop/NIS Saves/")
        default_filename = os.path.join(default_dir, filename)
        filename_dir = self.save_file_dialog(default_filename)

        # Check if filename is empty
        if filename_dir == "" or filename_dir is None:
            return

        # Get all the heading info from UI and update report and itinerary objects
        self.get_heading()

        # Set the type of document
        if save_type == "report":
            new_doc = self.itinerary.report.update_report_document(filename_dir)
        elif save_type == "itinerary":
            new_doc = self.itinerary.update_itinerary_document(filename_dir)
        else:
            return

        # Check if file is open
        if new_doc is True:
            self.show_popup("Saved", "Your file was successfully saved.", "info")
        else:
            self.show_popup("File Save Error", "The word document is currently open. Please close the file.", "error")

        # Set json date to the date of the Word document
        json_filedate = filename_dir[-16:].replace('.docx', '')
        # Check if the date is of the correct format
        try:
            datetime.datetime.strptime(json_filedate, "%d-%b-%Y")
        except ValueError:
            json_filedate = self.reformat_date(self.period_start)

        # Get the file name from the itinerary name
        filename_split = filename_dir.split("/")

        json_filedir = filename_dir.replace(filename_split[-1], '')
        json_filedir += f"state_saves"
        json_filename = f"NISIOB_{json_filedate}_{save_type}.json"

        self.save_to_json(json_filedir, json_filename)

    def save_to_json(self, file_dir, filename):
        # Add all the object info to dictionary
        itinerary = self.itinerary
        report = self.itinerary.report
        data = {
            "name": itinerary.name, "grade": itinerary.grade, "period_start": itinerary.period_start,
            "period_end": itinerary.period_end, "area": itinerary.area,
            "parish_office": itinerary.parish_office, "week_ended": report.week_ended,
            "files_carried_fwd": report.files_carried_fwd, "files_received": report.files_received,
            "total": report.total, "files_cleared": report.files_cleared,
            "files_on_hand": report.files_on_hand, "total_effective": report.total_effective,
            "total_ineffective": report.total_ineffective, "results_total": report.results_total,
            "objectives": itinerary.objectives
        }
        # Convert dictionary to json file
        json_data = json.dumps(data)

        # Save json file
        with open(file_dir+"/"+filename, "w") as file:
            file.write(json_data)

        # # Check if directory exist and create if necessary then save. -- Omit
        # try:
        #     os.mkdir(file_dir)
        # except OSError:
        #     with open(file_dir+"/"+filename, "w") as file:
        #         file.write(json_data)
        # else:
        #     with open(file_dir+"/"+filename, "w") as file:
        #         file.write(json_data)

    def autosave(self):
        self.create_files_folder()
        default_dir = os.path.expanduser("~/Desktop/NIS Saves/auto_saves")
        self.save_to_json(default_dir, f"autosave-{datetime.date.today()}.json")

    def save_file_dialog(self, default_filename):
        option = QFileDialog.Options()
        filename = QFileDialog.getSaveFileName(self, "Save word document", default_filename, "Word Document (*.docx)",
                                               options=option)
        return filename[0]

    @staticmethod
    def create_files_folder():
        default_dir = os.path.expanduser("~/Desktop/NIS Saves/")
        auto_saves_dir = os.path.join(default_dir, "auto_saves")
        state_saves_dir = os.path.join(default_dir, "state_saves")

        if os.path.exists(default_dir):
            pass
        else:
            os.mkdir(default_dir)

        # Create auto_saves_dir
        if os.path.exists(auto_saves_dir):
            pass
        else:
            os.mkdir(auto_saves_dir)

        # Create state_saves_dir
        if os.path.exists(state_saves_dir):
            pass
        else:
            os.mkdir(state_saves_dir)



# TODO #1 Update the results when loading - Done
# TODO #2 Create exception for json file error
# TODO #3 Allow edits to the itinerary - Partial
# TODO #4 Create tab to enter new businesses info
