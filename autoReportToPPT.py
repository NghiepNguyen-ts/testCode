## Import libraries
from copy_paste_to_ppt import *
from file_processing import *
from finalize_pptx import *


case_order = 1
list_folder = ['freesurface_left','freesurface_verPillar_left','freesurface_verPillar_right','freesurface_right','sp_ring2_lower_end_right','sp_ring2_lower_end_left','verPillar_lower_end_right','verPillar_lower_end_left', 'sp_ring_lower_end_right','sp_ring_lower_end_left']
list_folder_surface = ["Tank_inner_WallNoSlip1","tankSupport2_yz","tankSupport1_yz","midXZCrossSection"]
list_folder_total = ['freesurface_left','freesurface_verPillar_left','freesurface_verPillar_right','freesurface_right','sp_ring2_lower_end_right','sp_ring2_lower_end_left','verPillar_lower_end_right','verPillar_lower_end_left', 'sp_ring_lower_end_right','sp_ring_lower_end_left',"Tank_inner_WallNoSlip1","tankSupport2_yz","tankSupport1_yz","midXZCrossSection"]

for i in range(case_order,case_order+1):

    path_in_folder = "D:\\Yamanami\\Namura\\Rerun_202305\\Case" + str(i) + "_rerun_20230428\\postProcessing\\"
    ppt_path = "D:\\Yamanami\\Namura\\Rerun_202305\\Case" + str(i) + "_rerun_20230428\\Case" + str(i) + "_rerun_20230428.pptx"
    #path_source = "E:\\TechnoStar\\Project\\NAMURA\\SloshingTank\\Report\\Case" + str(i) + "_changeCenter"

    ## Run function 
    file_processing(list_folder,list_folder_surface,path_in_folder)
    #copy_file(list_folder,list_folder_surface, path_in_folder, path_source)
    copy_paste_to_PPT(list_folder,list_folder_surface,path_in_folder,ppt_path)
    finalize_ppt(ppt_path,list_folder_total,i,i)

