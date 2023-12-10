# coding: utf-8
from typing import List, Union

from PySide6.QtCore import Qt, QMargins, QModelIndex, QItemSelectionModel, Property,QRect, QEvent
from PySide6.QtGui import QPainter, QColor, QKeyEvent, QPalette, QBrush
from PySide6.QtWidgets import (QStyledItemDelegate, QApplication, QStyleOptionViewItem,
                               QTableView, QTableWidget, QWidget, QTableWidgetItem, QHeaderView)

from qfluentwidgets import getFont
from qfluentwidgets import isDarkTheme, FluentStyleSheet, themeColor
from qfluentwidgets import LineEdit, TextEdit, ComboBox
from qfluentwidgets import SmoothScrollDelegate

from typing import List, Set


class TableItemDelegate(QStyledItemDelegate):

    def __init__(self, parent: QTableView, col_indices: List[int] = None):
        super().__init__(parent)
        self.margin = 2
        self.hoverRow = -1
        self.pressedRow = -1
        self.selectedRows = set()
        self.col_indices = col_indices if col_indices is not None else []
        self.current_index = QModelIndex()
        self.installEventFilter(self)  # Install the event filter



    def setHoverRow(self, row: int):
        self.hoverRow = row
        self.parent().viewport().update()  # Force redraw of the viewpor

    def setPressedRow(self, row: int):
        self.pressedRow = row

    def setSelectedRows(self, indexes: List[QModelIndex]):
        self.selectedRows.clear()
        for index in indexes:
            self.selectedRows.add(index.row())
            if index.row() == self.pressedRow:
                self.pressedRow = -1

    def sizeHint(self, option, index):
        # increase original sizeHint to accommodate space needed for border
        size = super().sizeHint(option, index)
        size = size.grownBy(QMargins(0, self.margin, 0, self.margin))
        return size

    def eventFilter(self, source, event):
        '''This method is called when an event occurs on the watched object used for comboboxes'''
        #print(f"Event type: {event.type()}, Source type: {type(source)}")
        if isinstance(source, ComboBox) and event.type() == QEvent.HoverEnter:
            table = self.parent()  # The parent should be the QTableView
            index = table.indexAt(source.pos())
            self.setHoverRow(index.row())
            return True
        return super().eventFilter(source, event)

    def createEditor(self, parent: QWidget, option: QStyleOptionViewItem, index: QModelIndex) -> QWidget:
        self.current_index = index

        # Check if a QComboBox is assigned to this column
        for widget_type, columns in self.widgets.items():
            if index.column() in columns and widget_type == ComboBox:
                editor = widget_type(parent)  # create a ComboBox
                editor.installEventFilter(self)  # install the event filter here
                # Additional ComboBox setup logic here, if needed
                return editor  # return the ComboBox as the editor
            # Fallback to the default editor creation logic for other columns/widget types
        editor = super().createEditor(parent, option, index)

        # Logic specific to TextEdit
        if isinstance(editor, TextEdit):
            textEdit = editor  # Renaming for clarity in the following code
            textEdit.setPlainText(option.text)
            textEdit.document().adjustSize()
            docHeight = textEdit.document().size().height()
            textEdit.resize(option.rect.width(), docHeight)
            yOffset = (option.rect.height() - docHeight) // 2
            newRect = QRect(option.rect.x(), option.rect.y() - yOffset, option.rect.width(), docHeight)
            textEdit.setGeometry(newRect)
        # Logic specific to LineEdit
        elif isinstance(editor, LineEdit):
            lineEdit = editor  # Renaming for clarity in the following code
            lineEdit.setProperty("transparent", False)
            lineEdit.setStyle(QApplication.style())
            lineEdit.setText(option.text)
            lineEdit.setClearButtonEnabled(True)
        # Logic specific to QComboBox
        elif isinstance(editor, ComboBox):
            comboBox = editor  # Renaming for clarity in the following code
            # Set style for ComboBox here
            combo_box_style = """
                        ComboBox {
                            border: none;
                            border-radius: 0px;
                            padding: 5px 31px 6px 11px;
                            color: black;
                            background-color: transparent;
                            text-align: left;
                        }
                        ComboBox:hover {
                            background-color: rgba(249, 249, 249, 0.5);
                        }
                        ComboBox:pressed, ComboBox:on {
                            border: 1px solid rgba(0, 0, 0, 0.073);
                            border-radius: 5px;
                            border-bottom: 1px solid rgba(0, 0, 0, 0.183);
                            background-color: rgba(255, 255, 255, 0.7);
                        }
                        ComboBox:disabled {
                            color: rgba(0, 0, 0, 0.36);
                            background: rgba(249, 249, 249, 0.3);
                            border: 1px solid rgba(0, 0, 0, 0.06);
                            border-bottom: 1px solid rgba(0, 0, 0, 0.06);
                        }
                        """
            comboBox.setStyleSheet(combo_box_style)
        # Default logic
        else:
            # Creating a new LineEdit if editor is neither TextEdit nor LineEdit
            editor = LineEdit(parent)
            editor.setProperty("transparent", False)
            editor.setStyle(QApplication.style())
            editor.setText(option.text)
            editor.setClearButtonEnabled(True)

        return editor


    def updateEditorGeometry(self, editor: QWidget, option: QStyleOptionViewItem, index: QModelIndex):
        if isinstance(editor, TextEdit):  # Check if the editor is TextEdit
            editor.setGeometry(option.rect)
        if isinstance(editor, LineEdit):  # Check if the editor is TextEdit
            editor.setGeometry(option.rect)
            rect = option.rect
            y = rect.y() + (rect.height() - editor.height()) // 2
            x, w = max(8, rect.x()), rect.width()
            if index.column() == 0:
                w -= 8

            editor.setGeometry(x, y, w, rect.height())
        else:
            editor.setGeometry(option.rect)
            rect = option.rect
            y = rect.y() + (rect.height() - editor.height()) // 2
            x, w = max(8, rect.x()), rect.width()
            if index.column() == 0:
                w -= 8

            editor.setGeometry(x, y, w, rect.height())

    # Modify the _drawBackground method in the TableItemDelegate class
    def _drawBackground(self, painter: QPainter, option: QStyleOptionViewItem, index: QModelIndex):
        """ draw row background with rounded corners for the first and last visible column """
        r = 5
        model = index.model()
        total_columns = model.columnCount(index.parent())
        last_visible_column = total_columns - 1

        # Find the last visible column
        while last_visible_column > 0 and self.parent().isColumnHidden(last_visible_column):
            last_visible_column -= 1

        # Check for first column or last visible column
        if index.column() == 0 or index.column() == last_visible_column:
            # Adjust the rectangle for the rounded corners
            left_adjust = 4 if index.column() == 0 else -r - 1
            right_adjust = -4 if index.column() == last_visible_column else r + 1

            # Apply rounded corners
            rect = option.rect.adjusted(left_adjust, 0, right_adjust, 0)
            painter.drawRoundedRect(rect, r, r)
        else:
            # Draw regular rectangle for other columns
            rect = option.rect.adjusted(-1, 0, 1, 0)
            painter.drawRect(rect)

    def _drawIndicator(self, painter: QPainter, option: QStyleOptionViewItem, index: QModelIndex):
        """ draw indicator """
        y, h = option.rect.y(), option.rect.height()
        ph = round(0.35 * h if self.pressedRow == index.row() else 0.257 * h)
        painter.setBrush(themeColor())
        painter.drawRoundedRect(4, ph + y, 3, h - 2 * ph, 1.5, 1.5)

    def initStyleOption(self, option: QStyleOptionViewItem, index: QModelIndex):
        super().initStyleOption(option, index)

        # font
        option.font = index.data(Qt.FontRole) or getFont(13)

        # text color
        textColor = Qt.white if isDarkTheme() else Qt.black
        textBrush = index.data(Qt.ForegroundRole)  # type: QBrush
        if textBrush is not None:
            textColor = textBrush.color()

        option.palette.setColor(QPalette.Text, textColor)
        option.palette.setColor(QPalette.HighlightedText, textColor)

    def paint(self, painter, option, index):
        painter.save()
        painter.setPen(Qt.NoPen)
        painter.setRenderHint(QPainter.Antialiasing)

        # set clipping rect of painter to avoid painting outside the borders
        painter.setClipping(True)
        painter.setClipRect(option.rect)

        # call original paint method where option.rect is adjusted to account for border
        option.rect.adjust(0, self.margin, 0, -self.margin)

        # draw highlight background
        isHover = self.hoverRow == index.row()
        isPressed = self.pressedRow == index.row()
        isAlternate = index.row() % 2 == 0 and self.parent().alternatingRowColors()
        isDark = isDarkTheme()

        c = 255 if isDark else 0
        alpha = 0

        if index.row() not in self.selectedRows:
            if isPressed:
                alpha = 9 if isDark else 6
            elif isHover:
                alpha = 12
            elif isAlternate:
                alpha = 5
        else:
            if isPressed:
                alpha = 15 if isDark else 9
            elif isHover:
                alpha = 25
            else:
                alpha = 17

            # draw indicator
            if index.column() == 0 and self.parent().horizontalScrollBar().value() == 0:
                self._drawIndicator(painter, option, index)

        painter.setBrush(QColor(c, c, c, alpha))
        self._drawBackground(painter, option, index)

        # handle custom painting for ComboBox
        editor_type_to_handle = ComboBox  # Adjust this to the actual type you want to handle
        should_paint_custom = False

        for widget_type, columns in self.widgets.items():
            if index.column() in columns and widget_type == editor_type_to_handle:
                should_paint_custom = True
                break

        if should_paint_custom and index != self.current_index:
            # Only paint the text if the ComboBox editor is NOT open
            value = index.data(Qt.DisplayRole)
            text = value
            painter.drawText(option.rect, Qt.AlignLeft, text)
            painter.restore()
        else:
            # For all other cells, or when the ComboBox editor is open, delegate the painting to the base class
            painter.restore()
            super().paint(painter, option, index)


class TableBase:
    """ Table base class """

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.delegate = TableItemDelegate(self)
        self.scrollDelagate = SmoothScrollDelegate(self)
        self._isSelectRightClickedRow = False

        # set style sheet
        # FluentStyleSheet.TABLE_VIEW.apply(self)

        self.setStyleSheet("""
                            QTableView {
                                background: transparent;
                                outline: none;
                                border: none;
                                /* font: 13px 'Segoe UI', 'Microsoft YaHei'; */
                                selection-background-color: transparent;
                                alternate-background-color: transparent;
                            }

                            QTableView[isBorderVisible=true] {
                                border: 1px solid rgba(0, 0, 0, 15);
                            }

                            QTableView::item {
                                background: transparent;
                                border: 0px;
                                padding-left: 16px;
                                padding-right: 16px;
                                padding-top: 5px;  /* Add vertical padding */
                                padding-bottom: 5px;
                            }


                            QTableView::indicator {
                                width: 18px;
                                height: 18px;
                                border-radius: 5px;
                                border: 1px solid rgba(0, 0, 0, 0.48);
                                background-color: rgba(0, 0, 0, 0.022);
                            }

                            QTableView::indicator:hover {
                                border: 1px solid rgba(0, 0, 0, 0.56);
                                background-color: rgba(0, 0, 0, 0.05);
                            }

                            QTableView::indicator:pressed {
                                border: 1px solid rgba(0, 0, 0, 0.27);
                                background-color: rgba(0, 0, 0, 0.12);
                            }

                            QTableView::indicator:checked,
                            QTableView::indicator:indeterminate {
                                border: 1px solid --ThemeColorPrimary;
                                background-color: --ThemeColorPrimary;
                            }

                            QTableView::indicator:checked {
                                image: url(:/qfluentwidgets/images/check_box/Accept_white.svg);
                            }

                            QTableView::indicator:indeterminate {
                                image: url(:/qfluentwidgets/images/check_box/PartialAccept_white.svg);
                            }

                            QTableView::indicator:checked:hover,
                            QTableView::indicator:indeterminate:hover {
                                border: 1px solid --ThemeColorLight1;
                                background-color: --ThemeColorLight1;
                            }

                            QTableView::indicator:checked:pressed,
                            QTableView::indicator:indeterminate:pressed {
                                border: 1px solid --ThemeColorLight3;
                                background-color: --ThemeColorLight3;
                            }

                            QTableView::indicator:disabled {
                                border: 1px solid rgba(0, 0, 0, 0.27);
                                background-color: transparent;
                            }

                            QTableView::indicator:checked:disabled,
                            QTableView::indicator:indeterminate:disabled {
                                border: 1px solid rgb(199, 199, 199);
                                background-color: rgb(199, 199, 199);
                            }


                            QHeaderView {
                                background-color: transparent;
                            }

                            QHeaderView::section {
                                background-color: transparent;
                                color: rgb(96, 96, 96);
                                padding-left: 5px;
                                padding-right: 5px;
                                border: 1px solid rgba(0, 0, 0, 15);
                                font: 13px 'Segoe UI', 'Microsoft YaHei', 'PingFang SC';
                            }

                            QHeaderView::section:horizontal {
                                border-left: none;
                                height: 33px;
                            }

                            QTableView[isBorderVisible=true] QHeaderView::section:horizontal {
                                border-top: none;
                            }

                            QHeaderView::section:horizontal:last {
                                border-right: none;
                            }

                            QHeaderView::section:vertical {
                                border-top: none;
                                min-height: 20px; /* Adjust as needed */
                            }

                            QHeaderView::section:checked {
                                background-color: transparent;
                            }

                            QHeaderView::down-arrow {
                                subcontrol-origin: padding;
                                subcontrol-position: center right;
                                margin-right: 6px;
                                image: url(:/qfluentwidgets/images/table_view/Down_black.svg);
                            }

                            QHeaderView::up-arrow {
                                subcontrol-origin: padding;
                                subcontrol-position: center right;
                                margin-right: 6px;
                                image: url(:/qfluentwidgets/images/table_view/Up_black.svg);
                            }

                            QTableCornerButton::section {
                                background-color: transparent;
                                border: 1px solid rgba(0, 0, 0, 15);
                            }

                            QTableCornerButton::section:pressed {
                                background-color: rgba(0, 0, 0, 12);
                            }
                        """)

        self.setShowGrid(False)
        self.setMouseTracking(True)
        self.setAlternatingRowColors(True)
        self.setItemDelegate(self.delegate)
        self.setSelectionBehavior(TableWidget.SelectRows)
        self.horizontalHeader().setHighlightSections(False)
        self.verticalHeader().setHighlightSections(False)

        self.entered.connect(lambda i: self._setHoverRow(i.row()))
        self.pressed.connect(lambda i: self._setPressedRow(i.row()))
        self.verticalHeader().sectionClicked.connect(self.selectRow)


    def showEvent(self, e):
        QTableView.showEvent(self, e)
        self.resizeRowsToContents()

    def setBorderVisible(self, isVisible: bool):
        """ set the visibility of border """
        self.setProperty("isBorderVisible", isVisible)
        self.setStyle(QApplication.style())

    #def setBorderRadius(self, radius: int):
     #   """ set the radius of border """
     #   qss = f"QTableView{{border-radius: {radius}px}}"
     #   setCustomStyleSheet(self, qss, qss)

    def _setHoverRow(self, row: int):
        """ set hovered row """
        self.delegate.setHoverRow(row)
        self.viewport().update()

    def _setPressedRow(self, row: int):
        """ set pressed row """
        self.delegate.setPressedRow(row)
        self.viewport().update()

    def _setSelectedRows(self, indexes: List[QModelIndex]):
        self.delegate.setSelectedRows(indexes)
        self.viewport().update()

    def leaveEvent(self, e):
        QTableView.leaveEvent(self, e)
        self._setHoverRow(-1)

    def resizeEvent(self, e):
        QTableView.resizeEvent(self, e)
        self.viewport().update()

    def keyPressEvent(self, e: QKeyEvent):
        QTableView.keyPressEvent(self, e)
        self.updateSelectedRows()

    def mousePressEvent(self, e):
        if e.button() == Qt.LeftButton or self._isSelectRightClickedRow:
            return QTableView.mousePressEvent(self, e)

        index = self.indexAt(e.pos())
        if index.isValid():
            self._setPressedRow(index.row())

        QWidget.mousePressEvent(self, e)

    def mouseReleaseEvent(self, e):
        QTableView.mouseReleaseEvent(self, e)
        self.updateSelectedRows()

        if self.indexAt(e.pos()).row() < 0 or e.button() == Qt.RightButton:
            self._setPressedRow(-1)

    def setItemDelegate(self, delegate: TableItemDelegate):
        self.delegate = delegate
        super().setItemDelegate(delegate)

    def selectAll(self):
        QTableView.selectAll(self)
        self.updateSelectedRows()

    def selectRow(self, row: int):
        QTableView.selectRow(self, row)
        self.updateSelectedRows()

    def clearSelection(self):
        QTableView.clearSelection(self)
        self.updateSelectedRows()

    def setCurrentIndex(self, index: QModelIndex):
        QTableView.setCurrentIndex(self, index)
        self.updateSelectedRows()

    def updateSelectedRows(self):
        self._setSelectedRows(self.selectedIndexes())


class TableWidget(TableBase, QTableWidget):
    """ Table widget """

    def __init__(self, parent=None):
        super().__init__(parent)

    def setCurrentCell(self, row: int, column: int, command=None):
        self.setCurrentItem(self.item(row, column), command)

    def setCurrentItem(self, item: QTableWidgetItem, command=None):
        if not command:
            super().setCurrentItem(item)
        else:
            super().setCurrentItem(item, command)

        self.updateSelectedRows()

    def setCurrentCell(self, row: int, column: int, command=None):
        self.setCurrentItem(self.item(row, column), command)

    def setCurrentItem(self, item: QTableWidgetItem, command=None):
        if not command:
            super().setCurrentItem(item)
        else:
            super().setCurrentItem(item, command)

        self.updateSelectedRows()

    def isSelectRightClickedRow(self):
        return self._isSelectRightClickedRow

    def setSelectRightClickedRow(self, isSelect: bool):
        self._isSelectRightClickedRow = isSelect

    selectRightClickedRow = Property(bool, isSelectRightClickedRow, setSelectRightClickedRow)


class TableView(TableBase, QTableView):
    """ Table view """

    def __init__(self, parent=None):
        super().__init__(parent)

    def isSelectRightClickedRow(self):
        return self._isSelectRightClickedRow

    def setSelectRightClickedRow(self, isSelect: bool):
        self._isSelectRightClickedRow = isSelect

    selectRightClickedRow = Property(bool, isSelectRightClickedRow, setSelectRightClickedRow)

    # Override the paintEvent to use your custom drawing method
    def paintEvent(self, event):
        painter = QPainter(self.viewport())
        for row in range(self.model().rowCount()):
            for col in range(self.model().columnCount()):
                if not self.isColumnHidden(col):
                    index = self.model().index(row, col)
                    option = self.viewOptions()
                    option.rect = self.visualRect(index)
                    self.itemDelegate().paint(painter, option, index)
