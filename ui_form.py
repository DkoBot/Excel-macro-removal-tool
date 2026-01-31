# -*- coding: utf-8 -*-

################################################################################
## Form generated from reading UI file 'form.ui'
##
## Created by: Qt User Interface Compiler version 6.10.0
##
## WARNING! All changes made in this file will be lost when recompiling UI file!
################################################################################

from PySide6.QtCore import (QCoreApplication, QDate, QDateTime, QLocale,
    QMetaObject, QObject, QPoint, QRect,
    QSize, QTime, QUrl, Qt)
from PySide6.QtGui import (QBrush, QColor, QConicalGradient, QCursor,
    QFont, QFontDatabase, QGradient, QIcon,
    QImage, QKeySequence, QLinearGradient, QPainter,
    QPalette, QPixmap, QRadialGradient, QTransform)
from PySide6.QtWidgets import (QAbstractItemView, QApplication, QCheckBox, QComboBox,
    QHBoxLayout, QHeaderView, QLabel, QProgressBar,
    QPushButton, QSizePolicy, QSpacerItem, QTextEdit,
    QTreeWidget, QTreeWidgetItem, QVBoxLayout, QWidget)
import rc_Ico

class Ui_Widget(object):
    def setupUi(self, Widget):
        if not Widget.objectName():
            Widget.setObjectName(u"Widget")
        Widget.resize(430, 530)
        Widget.setMinimumSize(QSize(430, 530))
        Widget.setMaximumSize(QSize(582, 800))
        icon = QIcon()
        icon.addFile(u":/icopng/ico/\u6c47\u603b\u89e3\u9501.png", QSize(), QIcon.Mode.Normal, QIcon.State.Off)
        Widget.setWindowIcon(icon)
        self.verticalLayout = QVBoxLayout(Widget)
        self.verticalLayout.setObjectName(u"verticalLayout")
        self.horizontalLayout = QHBoxLayout()
        self.horizontalLayout.setObjectName(u"horizontalLayout")
        self.textEdit_2 = QTextEdit(Widget)
        self.textEdit_2.setObjectName(u"textEdit_2")
        self.textEdit_2.setMaximumSize(QSize(16777215, 25))
        self.textEdit_2.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.textEdit_2.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.textEdit_2.setLineWrapMode(QTextEdit.LineWrapMode.NoWrap)

        self.horizontalLayout.addWidget(self.textEdit_2)

        self.pushButton_2 = QPushButton(Widget)
        self.pushButton_2.setObjectName(u"pushButton_2")
        self.pushButton_2.setMaximumSize(QSize(16777215, 50))

        self.horizontalLayout.addWidget(self.pushButton_2)


        self.verticalLayout.addLayout(self.horizontalLayout)

        self.treeWidget = QTreeWidget(Widget)
        self.treeWidget.setObjectName(u"treeWidget")
        self.treeWidget.setContextMenuPolicy(Qt.ContextMenuPolicy.NoContextMenu)
        self.treeWidget.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.treeWidget.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.treeWidget.setAutoScroll(False)
        self.treeWidget.setProperty(u"showDropIndicator", True)
        self.treeWidget.setVerticalScrollMode(QAbstractItemView.ScrollMode.ScrollPerItem)

        self.verticalLayout.addWidget(self.treeWidget)

        self.horizontalLayout_3 = QHBoxLayout()
        self.horizontalLayout_3.setObjectName(u"horizontalLayout_3")
        self.checkBox = QCheckBox(Widget)
        self.checkBox.setObjectName(u"checkBox")
        self.checkBox.setMaximumSize(QSize(16777215, 20))

        self.horizontalLayout_3.addWidget(self.checkBox)

        self.verticalSpacer_3 = QSpacerItem(25, 10, QSizePolicy.Policy.Minimum, QSizePolicy.Policy.Expanding)

        self.horizontalLayout_3.addItem(self.verticalSpacer_3)

        self.label_2 = QLabel(Widget)
        self.label_2.setObjectName(u"label_2")
        self.label_2.setEnabled(True)
        self.label_2.setMaximumSize(QSize(100, 16777215))

        self.horizontalLayout_3.addWidget(self.label_2)

        self.textEdit_3 = QTextEdit(Widget)
        self.textEdit_3.setObjectName(u"textEdit_3")
        self.textEdit_3.setEnabled(False)
        self.textEdit_3.setMaximumSize(QSize(150, 25))
        self.textEdit_3.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.textEdit_3.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)

        self.horizontalLayout_3.addWidget(self.textEdit_3)


        self.verticalLayout.addLayout(self.horizontalLayout_3)

        self.horizontalLayout_2 = QHBoxLayout()
        self.horizontalLayout_2.setObjectName(u"horizontalLayout_2")
        self.checkBox_2 = QCheckBox(Widget)
        self.checkBox_2.setObjectName(u"checkBox_2")

        self.horizontalLayout_2.addWidget(self.checkBox_2)

        self.verticalSpacer_2 = QSpacerItem(15, 10, QSizePolicy.Policy.Minimum, QSizePolicy.Policy.Expanding)

        self.horizontalLayout_2.addItem(self.verticalSpacer_2)

        self.comboBox = QComboBox(Widget)
        self.comboBox.addItem("")
        self.comboBox.addItem("")
        self.comboBox.setObjectName(u"comboBox")

        self.horizontalLayout_2.addWidget(self.comboBox)

        self.verticalSpacer = QSpacerItem(15, 10, QSizePolicy.Policy.Minimum, QSizePolicy.Policy.Expanding)

        self.horizontalLayout_2.addItem(self.verticalSpacer)

        self.label = QLabel(Widget)
        self.label.setObjectName(u"label")
        self.label.setMaximumSize(QSize(100, 16777215))

        self.horizontalLayout_2.addWidget(self.label)

        self.textEdit = QTextEdit(Widget)
        self.textEdit.setObjectName(u"textEdit")
        self.textEdit.setEnabled(False)
        self.textEdit.setMaximumSize(QSize(150, 25))
        self.textEdit.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.textEdit.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)

        self.horizontalLayout_2.addWidget(self.textEdit)


        self.verticalLayout.addLayout(self.horizontalLayout_2)

        self.pushButton = QPushButton(Widget)
        self.pushButton.setObjectName(u"pushButton")
        self.pushButton.setMaximumSize(QSize(16777215, 50))

        self.verticalLayout.addWidget(self.pushButton)

        self.progressBar = QProgressBar(Widget)
        self.progressBar.setObjectName(u"progressBar")
        self.progressBar.setEnabled(True)
        self.progressBar.setMinimumSize(QSize(0, 0))
        self.progressBar.setMaximumSize(QSize(1500, 20))
        self.progressBar.setSizeIncrement(QSize(0, 50))
        self.progressBar.setBaseSize(QSize(0, 50))
        self.progressBar.setValue(0)
        self.progressBar.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.progressBar.setTextVisible(False)
        self.progressBar.setOrientation(Qt.Orientation.Horizontal)
        self.progressBar.setInvertedAppearance(False)
        self.progressBar.setTextDirection(QProgressBar.Direction.TopToBottom)

        self.verticalLayout.addWidget(self.progressBar)

        self.verticalLayout.setStretch(1, 1)

        self.retranslateUi(Widget)

        QMetaObject.connectSlotsByName(Widget)
    # setupUi

    def retranslateUi(self, Widget):
        Widget.setWindowTitle(QCoreApplication.translate("Widget", u"Widget", None))
        self.pushButton_2.setText(QCoreApplication.translate("Widget", u"\u9009\u4e2d\u6587\u4ef6\u6216\u6587\u4ef6\u5939", None))
        ___qtreewidgetitem = self.treeWidget.headerItem()
        ___qtreewidgetitem.setText(1, QCoreApplication.translate("Widget", u"\u6587\u4ef6\u540d", None));
        ___qtreewidgetitem.setText(0, QCoreApplication.translate("Widget", u"ID", None));
        self.checkBox.setText(QCoreApplication.translate("Widget", u"\u63d0\u53d6\u5b8f\u811a\u672c\u7528\u4e8e\u53d6\u8bc1\u751f\u6210\u62a5\u544a", None))
        self.label_2.setText(QCoreApplication.translate("Widget", u"\u5b8f\u603b\u5927\u5c0f:", None))
        self.checkBox_2.setText(QCoreApplication.translate("Widget", u"\u4fdd\u7559\u539f\u6587\u4ef6", None))
        self.comboBox.setItemText(0, QCoreApplication.translate("Widget", u"Chinese", None))
        self.comboBox.setItemText(1, QCoreApplication.translate("Widget", u"English", None))

        self.label.setText(QCoreApplication.translate("Widget", u"\u5171\u8ba1\u5904\u7406:", None))
        self.pushButton.setText(QCoreApplication.translate("Widget", u"\u5f00\u59cb\u6e05\u7406", None))
    # retranslateUi

