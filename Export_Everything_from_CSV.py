from connect import *
import os
clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import OpenFileDialog, DialogResult

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

def ChangePatient_UPMC(patient_db, MRN):
    info_all = patient_db.QueryPatientInfo(Filter={"FirstName": MRN,"LastName":'PT'})
    # If it isn't, see if it's in the secondary database
    if not info_all:
        info_all = patient_db.QueryPatientInfo(Filter={"PatientID": MRN}, UseIndexService=True)
    info = []
    for info_temp in info_all:
        if info_temp['FirstName'] == MRN:
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
            rois_in_case = []
            for roi in case.PatientModel.RegionsOfInterest:
                rois_in_case.append(roi.Name)
            for exam in case.Examinations:
                go = False
                for roi in rois_in_case:
                    if case.PatientModel.StructureSets[exam.Name].RoiGeometries[roi].HasContours():
                        go = True
                        break
                if not go:
                    continue
                out_path_export = os.path.join(path,MRN,exam.Name) # case.CaseName,
                if not os.path.exists(out_path_export):
                    Export_Dicom(path=out_path_export,case=case,exam=exam)
            continue
            for plan in case.TreatmentPlans:
                for beamset in plan.BeamSets:
                    beam_path = os.path.join(path, MRN, case.CaseName, plan.Name, beamset.DicomPlanLabel)
                    if not os.path.exists(beam_path):
                        os.makedirs(beam_path)
                    try:
                        case.ScriptableDicomExport(ExportFolderPath=beam_path,
                                                   BeamSetDoseForBeamSets=[beamset.BeamSetIdentifier()],
                                                   IgnorePreConditionWarnings=True)
                    except:
                        xxx = 1
                    print(1)
                    try:
                        case.ScriptableDicomExport(ExportFolderPath=beam_path,
                                                   BeamSets=[beamset.BeamSetIdentifier()],
                                                   IgnorePreConditionWarnings=True)
                    except:
                        xxx = 1

                    # print(3)
                    try:
                        case.ScriptableDicomExport(ExportFolderPath=beam_path,
                                                   TreatmentBeamDrrImages=[beamset.BeamSetIdentifier()])
                    except:
                        xxx = 1
                    print(4)
                    try:
                        case.ScriptableDicomExport(ExportFolderPath=beam_path,
                                                   SetupBeamDrrImages=[beamset.BeamSetIdentifier()],
                                                   IgnorePreConditionWarnings=True)
                    except:
                        xxx = 1
                    print(5)
                    try:
                        case.ScriptableDicomExport(ExportFolderPath=beam_path,
                                                   RtStructureSetsReferencedFromBeamSets=[beamset.BeamSetIdentifier()],
                                                   IgnorePreConditionWarnings=True)
                    except:
                        xxx = 1
            for for_registration in case.Registrations:
                to_for = for_registration.ToFrameOfReference
                # Frame of reference of the "From" examination.
                from_for = for_registration.FromFrameOfReference
                # Find all examinations with frame of reference that matches 'to_for'.
                to_examinations = [e for e in case.Examinations if
                                   e.EquipmentInfo.FrameOfReference == to_for]
        # Find all examinations with frame of reference that matches 'from_for'.
                from_examinations = [e for e in case.Examinations if e.EquipmentInfo.FrameOfReference == from_for]
                exam_names = ["%s:%s" % (for_registration.RegistrationSource.ToExamination.Name, for_registration.RegistrationSource.FromExamination.Name)]
                out_path = os.path.join(path,MRN,case.CaseName,'Registration', for_registration.RegistrationSource.FromExamination.Name)
                if not os.path.exists(out_path):
                    os.makedirs(out_path)
                try:
                    case.ScriptableDicomExport(ExportFolderPath=out_path,
                                               SpatialRegistrationForExaminations=exam_names,
                                               IgnorePreConditionWarnings=True)
                except:
                    continue


def main():
    pass


if __name__ == '__main__':
    main()
