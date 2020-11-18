__author__ = 'Brian M Anderson'
# Created on 11/18/2020
from connect import *
import os
clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import OpenFileDialog, DialogResult
'''
This is code to export out DICOM images and RT Structures

You need to provide a .csv file that has multiple columns
MRN Exam1, Exam2, Exam3, etc.
'''


def ChangePatient(patient_db, MRN):
    info_all = patient_db.QueryPatientInfo(Filter={"PatientID": MRN})
    # If it isn't, see if it's in the secondary database
    if not info_all:
        info_all = patient_db.QueryPatientInfo(Filter={"PatientID": MRN}, UseIndexService=True)
    info = []
    for info_temp in info_all:
        last_name = info_temp['LastName'].split('_')[-1]
        is_copy = False
        if last_name:
            for i in range(5):
                if last_name == str(i):
                    is_copy = True
                    break
        if info_temp['PatientID'] == MRN and not is_copy:
            info = info_temp
            break
    # If it isn't, see if it's in the secondary database
    patient = patient_db.LoadPatient(PatientInfo=info, AllowPatientUpgrade=True)
    return patient


def Export_Dicom(path,case,exam):
    if not os.path.exists(path):
        print('making path')
        os.makedirs(path)
    if not os.path.exists(os.path.join(path,'Completed.txt')): # If it has been previously uploaded, don't do it again
        case.ScriptableDicomExport(ExportFolderPath=path,Examinations=[], RtStructureSetsForExaminations=[exam.Name])
        case.ScriptableDicomExport(ExportFolderPath=path, Examinations=[exam.Name],
                                   RtStructureSetsForExaminations=[])
    return None


def get_file():
    dialog = OpenFileDialog()
    dialog.Filter = "All Files|*.*"
    result = dialog.ShowDialog()
    Path_File = ''
    if result == DialogResult.OK:
        Path_File = dialog.FileName
    if Path_File:
        return Path_File
    else:
        return None


def Export_to_path(path):
    patient_db = get_current('PatientDB')
    file = get_file()
    if file:
        fid = open(file)
    else:
        print('File not selected')
        return None
    data = {}
    for line in fid:
        line = line.strip('\n')
        line = line.split(',')
        if line and str(int(line[0])) == str(line[0]):
            data[line[0]] = [i.strip('"') for i in line[1:] if i != '']
    fid.close()
    if not os.path.exists(path):
        print('Path:' + path + ' does not exist')
        return None

    for MRN in data:
        print(MRN)
        # while len(MRN) < 3:
        #     MRN = '0' + MRN
        try:
            patient = ChangePatient(patient_db,MRN)
        except:
            continue
        '''
        Get all of the exams from the cases
        '''
        for case in patient.Cases:
            for exam in case.Examinations:
                if exam.Name in data[MRN]:
                    out_path_export = os.path.join(path, MRN, case.CaseName, exam.Name)
                    if not os.path.exists(out_path_export):
                        Export_Dicom(path=out_path_export, case=case, exam=exam)


def main():
    export_path = r'Path_to_Export'
    Export_to_path(export_path)


if __name__ == '__main__':
    main()
