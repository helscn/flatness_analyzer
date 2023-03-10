from PyQt5.QtWidgets import QWidget, QVBoxLayout

from matplotlib.backends.backend_qt5agg import FigureCanvas
from matplotlib.figure import Figure
from matplotlib.backends.backend_qt5agg import (
    NavigationToolbar2QT as NavigationToolbar)
from mpl_toolkits.mplot3d import Axes3D


class MplWidget(QWidget):

    def __init__(self, parent=None):
        super().__init__(parent)

        self.figure = Figure()
        self.canvas = FigureCanvas(self.figure)

        self.toolbar = NavigationToolbar(self.canvas, self)
        self.__changeActionLanguage()

        self.layout = QVBoxLayout(self)
        self.layout.addWidget(self.toolbar)
        self.layout.addWidget(self.canvas)

    def __changeActionLanguage(self):  # 汉化工具栏
        actList = self.toolbar.actions()  # 关联的Action列表
        actList[0].setText("复位")  # Home
        actList[0].setToolTip("复位到原始视图")  # Reset original view

        actList[1].setText("回退")  # Back
        actList[1].setToolTip("回退前一视图")  # Back to previous view

        actList[2].setText("前进")  # Forward
        actList[2].setToolTip("前进到下一视图")  # Forward to next view

        actList[4].setText("平动")  # Pan
        # Pan axes with left mouse, zoom with right
        actList[4].setToolTip("左键平移坐标轴，右键缩放坐标轴")

        actList[5].setText("缩放")  # Zoom
        actList[5].setToolTip("框选矩形框缩放")  # Zoom to rectangle

        actList[6].setText("外观")  # Subplots
        actList[6].setToolTip("外观设置")  # Configure subplots

        actList[7].setText("设置")  # Customize
        # Edit axis, curve and image parameters
        actList[7].setToolTip("定制图表参数")

        actList[9].setText("保存")  # Save
        actList[9].setToolTip("保存图表")  # Save the figure

        actList[-1].setVisible(True)
