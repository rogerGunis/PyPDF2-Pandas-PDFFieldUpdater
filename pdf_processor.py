#!/usr/bin/python3

import pandas as pd
import os
import random
from PyPDF2 import PdfWriter, PdfReader
from PyPDF2.generic import BooleanObject, NameObject, IndirectObject
from PyPDF2.generic import TextStringObject


def set_need_appearances_writer(writer: PdfWriter, reader: PdfReader):
    try:
        catalog = writer._root_object
        if "/AcroForm" not in catalog:
            reader_trailer = reader.trailer["/Root"]
            writer._root_object.update({
                key: reader_trailer[key]
                for key in reader_trailer
                if key in ("/AcroForm", "/Lang", "/MarkInfo")
                # NameObject("/AcroForm"): IndirectObject(len(writer._objects), 0, writer)
            })
        need_appearances = NameObject("/NeedAppearances")
        writer._root_object["/AcroForm"][need_appearances] = BooleanObject(True)
        
        return writer

    except Exception as e:
        print('set_need_appearances_writer() catch : ', repr(e))
        return writer

def updatePageFormFieldValues(page, fields):
        '''
        Update the form field values for a given page from a fields dictionary.

       This was copied from the PyPDF2 library and adapted for my use case.


        Copy field texts and values from fields to page.
        :param page: Page reference from PDF writer where the annotations
            and field data will be updated.
        :param fields: a Python dictionary of field names (/T) and text
            values (/V)
        '''
        # Iterate through pages, update field values
        for j in range(0, len(page['/Annots'])):
            writer_annot = page['/Annots'][j].get_object()
            field = writer_annot.get('/T') 
            if writer_annot.get("/FT") == "/Btn":
                value = fields.get(field, random.getrandbits(1))
                if value:
                    writer_annot.update(
                        {
                            NameObject("/AS"): NameObject("/On"),
                            NameObject("/V"): NameObject("/On"),
                        }
                    )
            elif writer_annot.get("/FT") == "/Tx":
                value = fields.get(field,field)
                writer_annot.update(
                    {
                        NameObject("/V"): TextStringObject(value),
                    }
                )

if __name__ == '__main__':
    csv_filename = "EISAutoFill.csv"
    pdf_filename = "test.pdf"
    
    csvin = os.path.normpath(os.path.join(os.getcwd(),'in',csv_filename))
    pdfin = os.path.normpath(os.path.join(os.getcwd(),'in',pdf_filename))
    pdfout = os.path.normpath(os.path.join(os.getcwd(),'out'))
    data = pd.read_csv(csvin)
    pdf = PdfReader(open(pdfin, "rb"), strict=False)  
    if "/AcroForm" in pdf.trailer["/Root"]:
        pdf.trailer["/Root"]["/AcroForm"].update(
            {NameObject("/NeedAppearances"): BooleanObject(True)})

    pdf_fields = [str(x) for x in pdf.get_fields().keys()] # List of all pdf field names
    # print(pdf_fields)
    csv_fields = data.columns.tolist()
    
    i = 0 #Filename numerical prefix
    for j, rows in data.iterrows():
        i += 1
        pdf_writer = PdfWriter()
        set_need_appearances_writer(pdf_writer, pdf)
        if "/AcroForm" in pdf_writer._root_object:
            pdf_writer._root_object["/AcroForm"].update(
                {NameObject("/NeedAppearances"): BooleanObject(True)})
        
        # Key = pdf_field_name : Value = csv_field_value
        field_dictionary_1 = {"S-03-001-TF": str(rows['FullName']),
                            "S-03-002-TF": rows['AddressLine1'],
                            "S-03-003-TF": rows['AddressLine2'],
                            "S-03-004-TF": rows['AddressLine3'],
                            "Post Code": rows['PostCode'],
                            "Description of Shares": rows['DescriptionOfShares'],
                            "Nominal Value of each Share": rows['NominalValueOfEachShare'],
                            "Number of Shares Issued": rows['NumberOfSharesIssued'],
                            "Amount Subscribed": rows['AmountSubscribed'],
                            "Share Issue Date": rows['ShareIssueDate'],
                            "Termination Date of these Shares": rows['TerminationDateOfTheseShares'],
                            "Received any value?": rows['ReceivedAnyValue?'],
                            "Name of Company Representative": rows['NameOfCompanyRepresentative'],
                            "Company Name": rows['CompanyName'],
                            "Unique Investment Reference Number": rows['UniqueInvestmentReferenceNumber'],
                            "Capacity in which signed": rows['CapacityInWhichSigned'],
                            "Registered Office Address Line 1": rows['RegisteredOfficeAddressLine1'],
                            "Registered Office Address Line 2": rows['RegisteredOfficeAddressLine2'],
                            "Registered Office Address Line 3": rows['RegisteredOfficeAddressLine3'],
                            "Date Signed": rows['DateSigned'],
                            "Post Code 2": rows['RegisteredOfficePostCode'],
                            }

        temp_out_dir = os.path.normpath(os.path.join(pdfout,str(i) + 'out.pdf'))
        pdf_writer.add_page(pdf.pages[3])
        updatePageFormFieldValues(page = pdf_writer.pages[0], fields = field_dictionary_1) # changed to user_defined_function
        pdf_writer.add_page(pdf.pages[4])
        pdf_writer.add_page(pdf.pages[5])
        pdf_writer.add_page(pdf.pages[6])
        outputStream = open(temp_out_dir, "wb")
        pdf_writer.write(outputStream)
        outputStream.close()
        print(f'Process Complete: {i} PDFs Processed!')
