import clr
import pyperclip
import os
import re
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from System.Diagnostics import Process
import glob
import xml.etree.ElementTree as ET
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")
from System.Drawing import SystemIcons
from System.Windows.Forms import (
    Application, Form, Label, TextBox, Button, OpenFileDialog, SaveFileDialog,
    DataGridView, DataGridViewAutoSizeColumnsMode, DataGridViewSelectionMode,
    DockStyle, GroupBox, NumericUpDown, DialogResult, FormStartPosition, FolderBrowserDialog, Keys, LinkLabel, AnchorStyles
)
from System.Drawing import Point, Size
from System import Decimal
from System.Windows.Forms import FolderBrowserDialog
from System.Windows.Forms import SaveFileDialog, DialogResult

def strip_formatting(s):
    if not isinstance(s, str):
        return s
    # Заменяем все последовательности вида \...; на пустое место
    return re.sub(r'\\[^;]*;', '', s).strip()

class MainForm(Form):
    def __init__(self):
        super().__init__()  
        self.Icon = SystemIcons.Application
                
        self.Text = "Просмотр и экспорт данных"
        self.Size = Size(900, 500)
        self.StartPosition = FormStartPosition.CenterScreen
        # ----- Path selection -----
        self.label_path = Label()
        self.label_path.Text = "Папка:"
        self.label_path.Location = Point(10, 10)
        self.label_path.AutoSize = True

        self.text_path = TextBox()
        self.text_path.Location = Point(70, 7)
        self.text_path.Width = 650

        self.btn_browse = Button()
        self.btn_browse.Text = "..."
        self.btn_browse.Location = Point(730, 5)
        self.btn_browse.Width = 30
        self.btn_browse.Click += self.browse_folder

        # ----- Data Table -----
        self.grid = DataGridView()
        self.grid.Location = Point(10, 40)
        self.grid.Size = Size(850, 200)
        self.grid.ReadOnly = True
        self.grid.AllowUserToAddRows = False
        self.grid.AllowUserToDeleteRows = False
        self.grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        self.grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect

        # ----- GroupBox Text 1 -----
        self.grp1 = GroupBox()
        self.grp1.Text = "Text 1"
        self.grp1.Location = Point(10, 250)
        self.grp1.Size = Size(420, 90)

        label1_minx = Label()
        label1_minx.Text = "minx"
        label1_minx.Location = Point(20, 32)
        label1_minx.AutoSize = True

        self.t1_minx = NumericUpDown()
        self.t1_minx.Location = Point(80, 28)
        self.t1_minx.Width = 80

        label1_maxx = Label()
        label1_maxx.Text = "maxx"
        label1_maxx.Location = Point(180, 32)
        label1_maxx.AutoSize = True

        self.t1_maxx = NumericUpDown()
        self.t1_maxx.Location = Point(240, 28)
        self.t1_maxx.Width = 80

        label1_miny = Label()
        label1_miny.Text = "miny"
        label1_miny.Location = Point(20, 62)
        label1_miny.AutoSize = True

        self.t1_miny = NumericUpDown()
        self.t1_miny.Location = Point(80, 58)
        self.t1_miny.Width = 80

        label1_maxy = Label()
        label1_maxy.Text = "maxy"
        label1_maxy.Location = Point(180, 62)
        label1_maxy.AutoSize = True

        self.t1_maxy = NumericUpDown()
        self.t1_maxy.Location = Point(240, 58)
        self.t1_maxy.Width = 80

        self.grp1.Controls.Add(label1_minx)
        self.grp1.Controls.Add(self.t1_minx)
        self.grp1.Controls.Add(label1_maxx)
        self.grp1.Controls.Add(self.t1_maxx)
        self.grp1.Controls.Add(label1_miny)
        self.grp1.Controls.Add(self.t1_miny)
        self.grp1.Controls.Add(label1_maxy)
        self.grp1.Controls.Add(self.t1_maxy)
        
        self.t1_maxx.Minimum = Decimal(-10000)
        self.t1_maxx.Maximum = Decimal(10000)
        self.t1_maxx.Value = Decimal(1000)

        self.t1_maxy.Minimum = Decimal(-10000)
        self.t1_maxy.Maximum = Decimal(10000)
        self.t1_maxy.Value = Decimal(1500)

        self.t1_minx.Minimum = Decimal(-10000)
        self.t1_minx.Maximum = Decimal(10000)
        self.t1_minx.Value = Decimal(0)

        self.t1_miny.Minimum = Decimal(-10000)
        self.t1_miny.Maximum = Decimal(10000)
        self.t1_miny.Value = Decimal(750)


        # ----- GroupBox Text 2 -----
        self.grp2 = GroupBox()
        self.grp2.Text = "Text 2"
        self.grp2.Location = Point(440, 250)
        self.grp2.Size = Size(420, 90)

        label2_minx = Label()
        label2_minx.Text = "minx"
        label2_minx.Location = Point(20, 32)
        label2_minx.AutoSize = True

        self.t2_minx = NumericUpDown()
        self.t2_minx.Location = Point(80, 28)
        self.t2_minx.Width = 80

        label2_maxx = Label()
        label2_maxx.Text = "maxx"
        label2_maxx.Location = Point(180, 32)
        label2_maxx.AutoSize = True

        self.t2_maxx = NumericUpDown()
        self.t2_maxx.Location = Point(240, 28)
        self.t2_maxx.Width = 80

        label2_miny = Label()
        label2_miny.Text = "miny"
        label2_miny.Location = Point(20, 62)
        label2_miny.AutoSize = True

        self.t2_miny = NumericUpDown()
        self.t2_miny.Location = Point(80, 58)
        self.t2_miny.Width = 80

        label2_maxy = Label()
        label2_maxy.Text = "maxy"
        label2_maxy.Location = Point(180, 62)
        label2_maxy.AutoSize = True

        self.t2_maxy = NumericUpDown()
        self.t2_maxy.Location = Point(240, 58)
        self.t2_maxy.Width = 80

        self.grp2.Controls.Add(label2_minx)
        self.grp2.Controls.Add(self.t2_minx)
        self.grp2.Controls.Add(label2_maxx)
        self.grp2.Controls.Add(self.t2_maxx)
        self.grp2.Controls.Add(label2_miny)
        self.grp2.Controls.Add(self.t2_miny)
        self.grp2.Controls.Add(label2_maxy)
        self.grp2.Controls.Add(self.t2_maxy)

        self.t2_maxx.Minimum = Decimal(-10000)
        self.t2_maxx.Maximum = Decimal(10000)
        self.t2_maxx.Value = Decimal(0)

        self.t2_maxy.Minimum = Decimal(-10000)
        self.t2_maxy.Maximum = Decimal(10000)
        self.t2_maxy.Value = Decimal(0)

        self.t2_minx.Minimum = Decimal(-10000)
        self.t2_minx.Maximum = Decimal(10000)
        self.t2_minx.Value = Decimal(0)

        self.t2_miny.Minimum = Decimal(-10000)
        self.t2_miny.Maximum = Decimal(10000)
        self.t2_miny.Value = Decimal(0)

        # ----- Buttons -----
        self.btn_refresh = Button()
        self.btn_refresh.Text = "Refresh"
        self.btn_refresh.Location = Point(690, 350)
        self.btn_refresh.Size = Size(80, 30)
        self.btn_refresh.Click += lambda sender, args: self.process_texts()
        
        self.btn_export = Button()
        self.btn_export.Text = "Export"
        self.btn_export.Location = Point(780, 400)
        self.btn_export.Size = Size(80, 30)
        self.btn_export.Click += self.export_data

        self.btn_cancel = Button()
        self.btn_cancel.Text = "Cancel"
        self.btn_cancel.Location = Point(690, 400)
        self.btn_cancel.Size = Size(80, 30)
        self.btn_cancel.Click += self.close_app

        # ----- Add controls -----
        self.Controls.Add(self.label_path)
        self.Controls.Add(self.text_path)
        self.Controls.Add(self.btn_browse)
        self.Controls.Add(self.grid)
        self.Controls.Add(self.grp1)
        self.Controls.Add(self.grp2)
        self.Controls.Add(self.btn_refresh)
        self.Controls.Add(self.btn_export)
        self.Controls.Add(self.btn_cancel)

        # ----- Data -----
        self.df = pd.DataFrame(columns=[
            "Имя файла", "имя блока", "minx", "maxx", "miny", "maxy", "text1", "text2"
        ])

        self.grid.KeyDown += self.grid_keydown
        
        # link to Github
        self.link = LinkLabel()
        self.link.Text = "Project: https://github.com/ariarchitect/izblock"
        self.link.Location = Point(10, self.ClientSize.Height - 40)  # 10px от левого края, 40px от нижнего
        self.link.AutoSize = True
        self.link.LinkClicked += self.link_clicked
        self.link.Anchor = AnchorStyles.Left | AnchorStyles.Bottom
        self.Controls.Add(self.link)

        self.load_demo_data()  # Replace with real data loading
        
    def link_clicked(self, sender, args):
        import System
        url = "https://github.com/ariarchitect/izblock"
        System.Diagnostics.Process.Start("explorer.exe", url)

    def browse_folder(self, sender, args):
        # тут вызываем свою функцию загрузки:
        folder_path = r"D:\Home\Downloads\tstsg;k;lgk"
        self.text_path.Text = folder_path
        self.load_izblock_xml(folder_path)
        process_texts()
        
    def process_texts(self):
        import xml.etree.ElementTree as ET
        import os
        
        # Диапазоны для группы 1
        t1_minx = float(str(self.t1_minx.Value))
        t1_maxx = float(str(self.t1_maxx.Value))
        t1_miny = float(str(self.t1_miny.Value))
        t1_maxy = float(str(self.t1_maxy.Value))
        
        # Диапазоны для группы 2
        t2_minx = float(str(self.t2_minx.Value))
        t2_maxx = float(str(self.t2_maxx.Value))
        t2_miny = float(str(self.t2_miny.Value))
        t2_maxy = float(str(self.t2_maxy.Value))
        
        
        for idx, row in self.df.iterrows():
            xml_file = os.path.join(self.text_path.Text, row['filename'])
            try:
                tree = ET.parse(xml_file)
                root = tree.getroot()
                for blk in root.findall('block'):
                    if blk.attrib.get('name') == row['name']:
                        # minx/miny блока — точка вставки
                        bx = float(blk.attrib.get('minx'))
                        by = float(blk.attrib.get('miny'))
                        t1_texts = []
                        t2_texts = []
                        for t in blk.findall('text'):
                            tx = float(t.attrib['x']) - bx
                            ty = float(t.attrib['y']) - by
                            content = t.attrib['content']
                            # для группы 1
                            if t1_minx <= tx <= t1_maxx and t1_miny <= ty <= t1_maxy:
                                t1_texts.append(content)
                            # для группы 2
                            if t2_minx <= tx <= t2_maxx and t2_miny <= ty <= t2_maxy:
                                t2_texts.append(content)
                        self.df.at[idx, 'text1'] = " ".join(t1_texts)
                        self.df.at[idx, 'text2'] = " ".join(t2_texts)
                        self.df['text1'] = self.df['text1'].apply(strip_formatting)
                        self.df['text2'] = self.df['text2'].apply(strip_formatting)
            except Exception as e:
                print("Ошибка парсинга или обработки блока", xml_file, e)
        self.refresh_table()

    def load_izblock_xml(self, folder_path):
        # --- тут функция загрузки xml, как выше ---
        import glob
        import xml.etree.ElementTree as ET
        import os
        import pandas as pd
        pattern = os.path.join(folder_path, "IZBLOCK_*.xml")
        files = glob.glob(pattern)
        blocks = []

        for file in files:
            try:
                tree = ET.parse(file)
                root = tree.getroot()
                for blk in root.findall('block'):
                    block_info = blk.attrib.copy()
                    block_info['filename'] = os.path.basename(file)
                    block_info['text1'] = ""
                    block_info['text2'] = ""
                    blocks.append(block_info)
            except Exception as e:
                print("Ошибка парсинга", file, e)
        columns = [
            "filename", "name", "minx", "maxx", "miny", "maxy", "text1", "text2"
        ]
        df = pd.DataFrame(blocks, columns=columns)
        self.df = df
        # Сразу после загрузки — вызываем обработку текстов:
        self.process_texts()
        self.refresh_table()

    def refresh_table(self):
        self.grid.Rows.Clear()
        self.grid.Columns.Clear()
        columns = [
            "filename", "name", "minx", "maxx", "miny", "maxy", "text1", "text2"
        ]
        for col in columns:
            self.grid.Columns.Add(col, col)
        for row in self.df.itertuples(index=False):
            self.grid.Rows.Add([str(x) if x is not None else "" for x in row])

    def browse_folder(self, sender, args):
        # Заглушка, путь фиксированный
        folder_path = r"D:\Home\Downloads\tstsg;k;lgk"
        self.text_path.Text = folder_path
        self.load_izblock_xml(folder_path)
        
    def export_data(self, sender, args):
        dialog = SaveFileDialog()
        dialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*"
        dialog.Title = "Export to Excel"
        dialog.FileName = "izblock_export.xlsx"
        if dialog.ShowDialog() == DialogResult.OK:
            filepath = dialog.FileName
            try:
                self.df.to_excel(filepath, index=False)
                from System.Windows.Forms import MessageBox
                MessageBox.Show("Export completed:\n" + filepath, "Export", 0)
            except Exception as e:
                MessageBox.Show("Export error:\n" + str(e), "Export", 0)



    def load_demo_data(self):
        # Заглушка! Заменим на реальную загрузку
        data = [
            ["file1.dwg", "BlockA", 1, 2, 3, 4, "text1A", "text2A"],
            ["file2.dwg", "BlockB", 5, 6, 7, 8, "text1B", "text2B"],
            ["file3.dwg", "BlockC", 9, 10, 11, 12, "text1C", "text2C"],
        ]
        self.df = pd.DataFrame(data, columns=self.df.columns)
        self.show_data_in_grid()

    def show_data_in_grid(self):
        self.grid.Rows.Clear()
        self.grid.Columns.Clear()
        for col in self.df.columns:
            self.grid.Columns.Add(col, col)
        for row in self.df.itertuples(index=False):
            self.grid.Rows.Add([str(x) for x in row])

    def close_app(self, sender, args):
        self.Close()
        
    def grid_keydown(self, sender, event):
        if event.Control and event.KeyCode == Keys.C:
            # Собираем выбранные строки и колонки
            rows = self.grid.SelectedRows
            cols = self.grid.SelectedCells
            if rows.Count > 0:
                # Если выбраны строки, копируем их все
                data = ""
                for row in rows:
                    vals = []
                    for cell in row.Cells:
                        vals.append(str(cell.Value))
                    data += "\t".join(vals) + "\n"
                pyperclip.copy(data)
            elif cols.Count > 0:
                # Если выбраны отдельные ячейки, копируем их значения
                data = ""
                for cell in cols:
                    data += str(cell.Value) + "\t"
                pyperclip.copy(data)
            # Можно показать всплывающее сообщение, если хочешь
            # MessageBox.Show("Скопировано!")

            # Отмечаем, что обработали событие
            event.Handled = True


if __name__ == "__main__":
    Application.EnableVisualStyles()
    form = MainForm()
    Application.Run(form)
    # После выхода из Application.Run — явно завершить процесс:
    import sys
    sys.exit()