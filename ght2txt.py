__author__ = "Da young Lee"
__credits__ = ["Seung hyup Hyun", "Jong chul Han", "Da young Lee"]
__version__ = "1.0.0"
__date__ = "2018.05.04"
__maintainer__ = "Da young Lee"
__email__ = "dyan.lee717@gmail.com"
__status__ = "Unofficial"

import os
import glob
import pandas as pd
import PyPDF2 as PDF
from PyPDF2 import PdfReader, PdfFileWriter


def mkdir_path(folder_dir):
    if not os.path.exists(folder_dir):
        os.mkdir(folder_dir)
    return folder_dir


class GHT2PDFconverter:
    def __init__(self, pdf_files, txt_files, cropped_dir):
        self.pdf_files = pdf_files
        self.txt_files = txt_files
        self.cropped_dir = cropped_dir
        self.inform = None


    def convert_pdf_to_txt(self):
        for pdf in self.pdf_files:
            reader = PdfReader(pdf)
            pages = reader.pages
            text = ""
            for page in pages:
                sub = page.extract_text()
                text += sub
            print(text)
            with open(pdf[:-4]+'.txt', 'w') as f:
                f.write(text)


    def store_information(self):
        informations = []
        for t, txt_file in enumerate(self.txt_files):
            f = open(txt_file, 'r')
            txt = f.read()
            split = (txt.split('\n'))
            info_ = []
            for s, split in enumerate(split):
                if s % 2 == 1:
                    info_.append(split)
            informations.append(info_)
        return informations


    def sort_information(self, informations):
        print('》 Storing informations...  \n')
        Values, Quadras = [], []
        for i, self.inform in enumerate(informations):
            Index = str('%04d' % (i+1))  # .zfill(3)
            FileName = self.txt_files[i][6:-4]+'.jpg'
            Patient = self.inform[0][9:-2]
            PatientID = self.inform[3][12:].rjust(5, '0')
            ODOS = self.inform[4][:3]
            
            for j, info_line in enumerate(self.inform):
                if info_line == 'Date:':
                    TDate = self.inform[j+3][:]
                    Age = self.inform[j+4][-3:-1]
                if info_line.split(':')[0] == 'GHT':
                    info_below = self.inform[j+1].split(':')
                    if info_line.split(':')[1][-3:] == 'VFI':
                        GHT = self.inform[j].split(':')[1][:-3]
                        VFI = info_line.split(':')[2][:-1]
                        MD = info_below[1][:6]
                        PSD = info_below[2][:6]
                    elif info_below[0] == 'MD':
                        GHT = self.inform[j].split(':')[1]
                        MD = info_below[1][:6]
                        PSD = info_below[2][:6]
                        VFI = info_below[3][:-2]
                    elif info_below[0] == 'VFI':
                        GHT = self.inform[j].split(':')[1]
                        VFI = info_below[1][:-1]
                        info_below2 = self.inform[j+2].split(':')
                        MD = info_below2[1][:6]
                        PSD = info_below2[2][:6]
                    else:
                        print('wrong!!!', i, PatientID)

                if len(self.inform) >= 100 and info_line[:5] == 'Total':
                    if ODOS == 'OD ':
                        Q2 = self.index_adjustment(j-36, j-30, 12, -1)
                        Q1 = self.index_adjustment(j-30, j-24, 10, -1, True, 16)
                        Q3 = self.index_adjustment(j-12, j-6, 10, 0)
                        Q4 = self.index_adjustment(j-6, j, 6, 0, True, 3)
                        Quadras.append([Q1, Q2, Q3, Q4])
                    else:
                        Q2 = self.index_adjustment(j-36, j-30, 12, -1, True, 16)
                        Q1 = self.index_adjustment(j-30, j-24, 10, -1)
                        Q3 = self.index_adjustment(j-12, j-6, 8, 0, True, 3)
                        Q4 = self.index_adjustment(j-6, j, 7, 0)
                        Quadras.append([Q1, Q2, Q3, Q4])

            if len(self.inform) < 100:
                print('     [ Pattern Deviation Graph NOT Exists! ]')
                print('       Index=', i, ',  PatientID=', PatientID, ',  Lenght of text=', len(self.inform), '\n')
                Quadras += [[['  'for _ in range(19)]] * 4]

            Values.append([Index, FileName, Patient, PatientID, TDate, Age, ODOS, GHT, MD, PSD, VFI])
        
        return Values, Quadras

    
    def index_adjustment(self, range_a, range_b, insert_idx, apart_idx, blank=False, blk_idx=None):
        quadrant = self.inform[range_a:range_b]
        letters = []
        for l, line in enumerate(quadrant):
            letters.extend(quadrant[l].split())
        letters.insert(insert_idx, letters[apart_idx])
        if blank is True:
            letters.insert(blk_idx, ' ')
        del (letters[apart_idx])
        return letters

    
    def set_values(self, Values):
        # Dataframe for patient information 'Values'
        index_list = ['Index', 'FileName', 'Patient', 'PatientID', 'TDate', 'Age',
                      'OD/OS', 'GHT', 'MD(dB)', 'PSD(dB)', 'VFI(%)']
        df1 = pd.DataFrame(index=index_list)
        dataset = []
        for i, value in enumerate(Values):
            dataset = list(zip(index_list, value))
            mini_df = pd.DataFrame(data=dataset, index=index_list)
            df1 = pd.concat([df1, mini_df.iloc[:, [1]]], axis=1, ignore_index=True)
        df1 = df1.T
        return df1
    

    def set_quadras(self, Quadras):
        # Dataframe for diviation pattern graph 'Quadras'
        Quadra_index, Quadra_index_list = [], []
        Q1_index, Q2_index, Q3_index, Q4_index = [], [], [], []
        for i in range(19):
            Q1_index.append('Q1'+'_'+str(i+1))
            Q2_index.append('Q2'+'_'+str(i+1))
            Q3_index.append('Q3'+'_'+str(i+1))
            Q4_index.append('Q4'+'_'+str(i+1))
        Quadra_index.extend([Q1_index]+[Q2_index]+[Q3_index]+[Q4_index])  # index that will saved after cropping
        Quadra_index_list.extend(Q1_index+Q2_index+Q3_index+Q4_index)  # index that will saved entirelly
        df2 = pd.DataFrame(index=Quadra_index_list)
        for i, quadras in enumerate(Quadras):
            dataset = []
            for q, quadra in enumerate(quadras):
                dataset.extend(list(zip(Quadra_index[q], quadra)))
            mini_df = pd.DataFrame(data=dataset, index=Quadra_index_list)
            df2 = pd.concat([df2, mini_df.iloc[:, [1]]], axis=1, ignore_index=True)
        df2 = df2.T
        return df2


    def write_excel(self, writer):
        # Working with pandas excel writer to decorate cells
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        header_fmt = workbook.add_format()  # ({'bold': True, 'border': 2})
        header_fmt.set_font_size(9)
        header_fmt.set_align('center')
        header_fmt.set_align('vcenter')
        worksheet.set_row(0, 17, header_fmt)
        cell_fmt = workbook.add_format()
        cell_fmt.set_font_size(9)
        cell_fmt.set_align('center')
        cell_fmt.set_align('vcenter')
        worksheet.set_column(0, 0, 4.5, cell_fmt)
        worksheet.set_column(1, 1, 12, cell_fmt)
        worksheet.set_column(2, 2, 17, cell_fmt)
        worksheet.set_column(3, 3, 9, cell_fmt)
        worksheet.set_column(4, 4, 11, cell_fmt)
        worksheet.set_column(5, 5, 4, cell_fmt)
        worksheet.set_column(6, 6, 6, cell_fmt)
        worksheet.set_column(7, 7, 21, cell_fmt)
        worksheet.set_column(8, 10, 7, cell_fmt)
        worksheet.set_column(11, 87, 5.5, cell_fmt)
        worksheet.set_column(87, 87, 10, cell_fmt)
        writer._save()
            
    def crop_GHT(self):
        for i, PDFfile in enumerate(self.pdf_files):
            input_ = PDF.PdfFileReader(open(PDFfile, 'rb'))
            output_ = PdfFileWriter()
            numPages = input_.getNumPages()  # They all has only 1 page.
            Page = input_.getPage(numPages-1)
            # UpperRight_x, UpperRight_y = Page.mediaBox.getUpperRight_x(), Page.mediaBox.getUpperRight_y()
            # (595.28, 841.89) ...and this means coordinate start with LowerLeft
            Page.trimBox.lowerLeft = (330, 439)
            Page.trimBox.upperRight = (513, 620)
            Page.cropBox.lowerLeft = (513, 439)
            Page.cropBox.upperRight = (330, 620)
            output_.addPage(Page)
            output_name = os.path.join(self.cropped_dir+self.pdf_files[i][5:-4]+"_"+self.pdf_files[i][-4:])
            outputStream = open(output_name, "wb")
            output_.write(outputStream)
            outputStream.close()


if __name__ == '__main__':
    
    data_dir = mkdir_path('ZEISS/')
    pdf_files = glob.glob(data_dir+'*.pdf')
    txt_files = glob.glob(data_dir+'*.txt')
    print('\n')
    print('》 PDF and text data directory is, "', data_dir, '"')
    print('》 There are  (', len(txt_files), ')  text files in data directory.  \n')
    out_dir = mkdir_path('output/')
    cropped_dir = mkdir_path(os.path.join(out_dir + 'cropped/'))
    
    ght2pdf = GHT2PDFconverter(pdf_files, txt_files, cropped_dir)
    # ght2pdf.convert_pdf_to_txt()
    informations = ght2pdf.store_information()
    Values, Quadras = ght2pdf.sort_information(informations)
    df1 = ght2pdf.set_values(Values)
    df2 = ght2pdf.set_quadras(Quadras)
        
    result = pd.merge(df1, df2, left_index=True, right_index=True, how='outer')
    # For futher analysis, add a column for grouping -> GHT_grouper.py
    result['group'] = pd.Series(list(' '*len(txt_files)))  # , index=result.index)
    writer = pd.ExcelWriter(os.path.join(out_dir+'result.xlsx'), engine='xlsxwriter')
    result.to_excel(writer, 'Sheet1', index=False)
    ght2pdf.write_excel(writer)
    print('》 Worksheet is saved!')
    print('》 Please check on excel file in the output directory. \n  The output directory is, "', out_dir, '" \n')

    print('》 Cropping PDF files...')
    ght2pdf.crop_GHT()
    print('》 Every PDF file is cropped. \n  The cropped image directory is, "', cropped_dir, '"')
    print('》 Please convert to JPG image. \n\n')
    
    pause = input('Press ENTER key to exit >>> ')
