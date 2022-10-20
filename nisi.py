from itinerary import Itinerary
from report import Report
import sys
from ui import *


if __name__ == "__main__":
    app = QApplication(sys.argv)

    report = Report()
    itinerary = Itinerary(report)
    main_window = MainWindow(itinerary)
    sys.exit(app.exec())
