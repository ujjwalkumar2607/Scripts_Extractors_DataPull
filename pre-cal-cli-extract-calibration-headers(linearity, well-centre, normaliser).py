import os
import glob
import tarfile
import gzip
import json
import pandas as pd

def process_tar_file(tar_file, output_df, row_index):
    
    with tarfile.open(tar_file, "r:*") as tar:
        hostname_data = None
        ref_unit_data = None
        batch_code_data = None
        date_time_data = None
        red_16x_w1 = None
        red_16x_w2 = None
        red_16x_w3 = None
        
        for member in tar.getmembers():
            if member.name.endswith("device_information.json"):
                f = tar.extractfile(member)
                json_data = json.load(f)
                hostname_data = json_data["hostname"]
            elif member.name.endswith("optical_cal_targets_gen_3_0.json"):
                f = tar.extractfile(member)
                json_data = json.load(f)
                ref_unit_data = json_data["ref_unit"]
                batch_code_data = json_data["batchcode"]
                date_time_data = json_data["timestamp"]
            elif member.name.endswith("well_center_store.json"):
                f = tar.extractfile(member)
                json_data = json.load(f)
                well_1_center = json_data['test_results']['wells']['well_1']
                well_2_center = json_data['test_results']['wells']['well_2']
                well_3_center = json_data['test_results']['wells']['well_3']
            elif member.name.endswith("optical_store_gen3_storage.json"):
                f = tar.extractfile(member)
                json_data = json.load(f)

                                #Red Standards 1,2,3
                red_16x_w1 = json_data['test_results']['analysis_values']['precal']['standard_1']['img1'][0]
                red_16x_w2 = json_data['test_results']['analysis_values']['precal']['standard_1']['img1'][1]
                red_16x_w3 = json_data['test_results']['analysis_values']['precal']['standard_1']['img1'][2]
                red_16x_w1_p2 = json_data['test_results']['analysis_values']['precal']['standard_1']['img2'][0]
                red_16x_w2_p2 = json_data['test_results']['analysis_values']['precal']['standard_1']['img2'][1]
                red_16x_w3_p2 = json_data['test_results']['analysis_values']['precal']['standard_1']['img2'][2]
                red_16x_w1_p3 = json_data['test_results']['analysis_values']['precal']['standard_1']['img3'][0]
                red_16x_w2_p3 = json_data['test_results']['analysis_values']['precal']['standard_1']['img3'][1]
                red_16x_w3_p3 = json_data['test_results']['analysis_values']['precal']['standard_1']['img3'][2]

                red_32x_w1 = json_data['test_results']['analysis_values']['precal']['standard_2']['img1'][0]
                red_32x_w2 = json_data['test_results']['analysis_values']['precal']['standard_2']['img1'][1]
                red_32x_w3 = json_data['test_results']['analysis_values']['precal']['standard_2']['img1'][2]
                red_32x_w1_p2 = json_data['test_results']['analysis_values']['precal']['standard_2']['img2'][0]
                red_32x_w2_p2 = json_data['test_results']['analysis_values']['precal']['standard_2']['img2'][1]
                red_32x_w3_p2 = json_data['test_results']['analysis_values']['precal']['standard_2']['img2'][2]
                red_32x_w1_p3 = json_data['test_results']['analysis_values']['precal']['standard_2']['img3'][0]
                red_32x_w2_p3 = json_data['test_results']['analysis_values']['precal']['standard_2']['img3'][1]
                red_32x_w3_p3 = json_data['test_results']['analysis_values']['precal']['standard_2']['img3'][2]

                red_64x_w1 = json_data['test_results']['analysis_values']['precal']['standard_3']['img1'][0]
                red_64x_w2 = json_data['test_results']['analysis_values']['precal']['standard_3']['img1'][1]
                red_64x_w3 = json_data['test_results']['analysis_values']['precal']['standard_3']['img1'][2]
                red_64x_w1_p2 = json_data['test_results']['analysis_values']['precal']['standard_3']['img2'][0]
                red_64x_w2_p2 = json_data['test_results']['analysis_values']['precal']['standard_3']['img2'][1]
                red_64x_w3_p2 = json_data['test_results']['analysis_values']['precal']['standard_3']['img2'][2]
                red_64x_w1_p3 = json_data['test_results']['analysis_values']['precal']['standard_3']['img3'][0]
                red_64x_w2_p3 = json_data['test_results']['analysis_values']['precal']['standard_3']['img3'][1]
                red_64x_w3_p3 = json_data['test_results']['analysis_values']['precal']['standard_3']['img3'][2]

                #Red Standards 4,5,6
                red_16x_w1_std4 = json_data['test_results']['analysis_values']['precal']['standard_4']['img1'][0]
                red_16x_w2_std4 = json_data['test_results']['analysis_values']['precal']['standard_4']['img1'][1]
                red_16x_w3_std4 = json_data['test_results']['analysis_values']['precal']['standard_4']['img1'][2]
                red_16x_w1_p2_std4 = json_data['test_results']['analysis_values']['precal']['standard_4']['img2'][0]
                red_16x_w2_p2_std4 = json_data['test_results']['analysis_values']['precal']['standard_4']['img2'][1]
                red_16x_w3_p2_std4 = json_data['test_results']['analysis_values']['precal']['standard_4']['img2'][2]
                red_16x_w1_p3_std4 = json_data['test_results']['analysis_values']['precal']['standard_4']['img3'][0]
                red_16x_w2_p3_std4 = json_data['test_results']['analysis_values']['precal']['standard_4']['img3'][1]
                red_16x_w3_p3_std4 = json_data['test_results']['analysis_values']['precal']['standard_4']['img3'][2]

                red_32x_w1_std5 = json_data['test_results']['analysis_values']['precal']['standard_5']['img1'][0]
                red_32x_w2_std5 = json_data['test_results']['analysis_values']['precal']['standard_5']['img1'][1]
                red_32x_w3_std5 = json_data['test_results']['analysis_values']['precal']['standard_5']['img1'][2]
                red_32x_w1_p2_std5 = json_data['test_results']['analysis_values']['precal']['standard_5']['img2'][0]
                red_32x_w2_p2_std5 = json_data['test_results']['analysis_values']['precal']['standard_5']['img2'][1]
                red_32x_w3_p2_std5 = json_data['test_results']['analysis_values']['precal']['standard_5']['img2'][2]
                red_32x_w1_p3_std5 = json_data['test_results']['analysis_values']['precal']['standard_5']['img3'][0]
                red_32x_w2_p3_std5 = json_data['test_results']['analysis_values']['precal']['standard_5']['img3'][1]
                red_32x_w3_p3_std5 = json_data['test_results']['analysis_values']['precal']['standard_5']['img3'][2]

                red_64x_w1_std6 = json_data['test_results']['analysis_values']['precal']['standard_6']['img1'][0]
                red_64x_w2_std6 = json_data['test_results']['analysis_values']['precal']['standard_6']['img1'][1]
                red_64x_w3_std6 = json_data['test_results']['analysis_values']['precal']['standard_6']['img1'][2]
                red_64x_w1_p2_std6 = json_data['test_results']['analysis_values']['precal']['standard_6']['img2'][0]
                red_64x_w2_p2_std6 = json_data['test_results']['analysis_values']['precal']['standard_6']['img2'][1]
                red_64x_w3_p2_std6 = json_data['test_results']['analysis_values']['precal']['standard_6']['img2'][2]
                red_64x_w1_p3_std6 = json_data['test_results']['analysis_values']['precal']['standard_6']['img3'][0]
                red_64x_w2_p3_std6 = json_data['test_results']['analysis_values']['precal']['standard_6']['img3'][1]
                red_64x_w3_p3_std6 = json_data['test_results']['analysis_values']['precal']['standard_6']['img3'][2]

                #Green Standards 1,2,3
                green_16x_w1 = json_data['test_results']['analysis_values']['precal']['standard_1']['img1'][3]
                green_16x_w2 = json_data['test_results']['analysis_values']['precal']['standard_1']['img1'][4]
                green_16x_w3 = json_data['test_results']['analysis_values']['precal']['standard_1']['img1'][5]
                green_16x_w1_p2 = json_data['test_results']['analysis_values']['precal']['standard_1']['img2'][3]
                green_16x_w2_p2 = json_data['test_results']['analysis_values']['precal']['standard_1']['img2'][4]
                green_16x_w3_p2 = json_data['test_results']['analysis_values']['precal']['standard_1']['img2'][5]
                green_16x_w1_p3 = json_data['test_results']['analysis_values']['precal']['standard_1']['img3'][3]
                green_16x_w2_p3 = json_data['test_results']['analysis_values']['precal']['standard_1']['img3'][4]
                green_16x_w3_p3 = json_data['test_results']['analysis_values']['precal']['standard_1']['img3'][5]

                green_32x_w1 = json_data['test_results']['analysis_values']['precal']['standard_2']['img1'][3]
                green_32x_w2 = json_data['test_results']['analysis_values']['precal']['standard_2']['img1'][4]
                green_32x_w3 = json_data['test_results']['analysis_values']['precal']['standard_2']['img1'][5]
                green_32x_w1_p2 = json_data['test_results']['analysis_values']['precal']['standard_2']['img2'][3]
                green_32x_w2_p2 = json_data['test_results']['analysis_values']['precal']['standard_2']['img2'][4]
                green_32x_w3_p2 = json_data['test_results']['analysis_values']['precal']['standard_2']['img2'][5]
                green_32x_w1_p3 = json_data['test_results']['analysis_values']['precal']['standard_2']['img3'][3]
                green_32x_w2_p3 = json_data['test_results']['analysis_values']['precal']['standard_2']['img3'][4]
                green_32x_w3_p3 = json_data['test_results']['analysis_values']['precal']['standard_2']['img3'][5]

                green_64x_w1 = json_data['test_results']['analysis_values']['precal']['standard_3']['img1'][3]
                green_64x_w2 = json_data['test_results']['analysis_values']['precal']['standard_3']['img1'][4]
                green_64x_w3 = json_data['test_results']['analysis_values']['precal']['standard_3']['img1'][5]
                green_64x_w1_p2 = json_data['test_results']['analysis_values']['precal']['standard_3']['img2'][3]
                green_64x_w2_p2 = json_data['test_results']['analysis_values']['precal']['standard_3']['img2'][4]
                green_64x_w3_p2 = json_data['test_results']['analysis_values']['precal']['standard_3']['img2'][5]
                green_64x_w1_p3 = json_data['test_results']['analysis_values']['precal']['standard_3']['img3'][3]
                green_64x_w2_p3 = json_data['test_results']['analysis_values']['precal']['standard_3']['img3'][4]
                green_64x_w3_p3 = json_data['test_results']['analysis_values']['precal']['standard_3']['img3'][5]

                #Green Standards 4,5,6
                green_16x_w1_std4 = json_data['test_results']['analysis_values']['precal']['standard_4']['img1'][3]
                green_16x_w2_std4 = json_data['test_results']['analysis_values']['precal']['standard_4']['img1'][4]
                green_16x_w3_std4 = json_data['test_results']['analysis_values']['precal']['standard_4']['img1'][5]
                green_16x_w1_p2_std4 = json_data['test_results']['analysis_values']['precal']['standard_4']['img2'][3]
                green_16x_w2_p2_std4 = json_data['test_results']['analysis_values']['precal']['standard_4']['img2'][4]
                green_16x_w3_p2_std4 = json_data['test_results']['analysis_values']['precal']['standard_4']['img2'][5]
                green_16x_w1_p3_std4 = json_data['test_results']['analysis_values']['precal']['standard_4']['img3'][3]
                green_16x_w2_p3_std4 = json_data['test_results']['analysis_values']['precal']['standard_4']['img3'][4]
                green_16x_w3_p3_std4 = json_data['test_results']['analysis_values']['precal']['standard_4']['img3'][5]

                green_32x_w1_std5 = json_data['test_results']['analysis_values']['precal']['standard_5']['img1'][3]
                green_32x_w2_std5 = json_data['test_results']['analysis_values']['precal']['standard_5']['img1'][4]
                green_32x_w3_std5 = json_data['test_results']['analysis_values']['precal']['standard_5']['img1'][5]
                green_32x_w1_p2_std5 = json_data['test_results']['analysis_values']['precal']['standard_5']['img2'][3]
                green_32x_w2_p2_std5 = json_data['test_results']['analysis_values']['precal']['standard_5']['img2'][4]
                green_32x_w3_p2_std5 = json_data['test_results']['analysis_values']['precal']['standard_5']['img2'][5]
                green_32x_w1_p3_std5 = json_data['test_results']['analysis_values']['precal']['standard_5']['img3'][3]
                green_32x_w2_p3_std5 = json_data['test_results']['analysis_values']['precal']['standard_5']['img3'][4]
                green_32x_w3_p3_std5 = json_data['test_results']['analysis_values']['precal']['standard_5']['img3'][5]

                green_64x_w1_std6 = json_data['test_results']['analysis_values']['precal']['standard_6']['img1'][3]
                green_64x_w2_std6 = json_data['test_results']['analysis_values']['precal']['standard_6']['img1'][4]
                green_64x_w3_std6 = json_data['test_results']['analysis_values']['precal']['standard_6']['img1'][5]
                green_64x_w1_p2_std6 = json_data['test_results']['analysis_values']['precal']['standard_6']['img2'][3]
                green_64x_w2_p2_std6 = json_data['test_results']['analysis_values']['precal']['standard_6']['img2'][4]
                green_64x_w3_p2_std6 = json_data['test_results']['analysis_values']['precal']['standard_6']['img2'][5]
                green_64x_w1_p3_std6 = json_data['test_results']['analysis_values']['precal']['standard_6']['img3'][3]
                green_64x_w2_p3_std6 = json_data['test_results']['analysis_values']['precal']['standard_6']['img3'][4]
                green_64x_w3_p3_std6 = json_data['test_results']['analysis_values']['precal']['standard_6']['img3'][5]

                #Buffer Red 7,8,9
                buffer_red_16x_w1 = json_data['test_results']['analysis_values']['precal']['standard_7']['img1'][0]
                buffer_red_16x_w2 = json_data['test_results']['analysis_values']['precal']['standard_7']['img1'][1]
                buffer_red_16x_w3 = json_data['test_results']['analysis_values']['precal']['standard_7']['img1'][2]
                buffer_red_16x_w1_p2 = json_data['test_results']['analysis_values']['precal']['standard_7']['img2'][0]
                buffer_red_16x_w2_p2 = json_data['test_results']['analysis_values']['precal']['standard_7']['img2'][1]
                buffer_red_16x_w3_p2 = json_data['test_results']['analysis_values']['precal']['standard_7']['img2'][2]
                buffer_red_16x_w1_p3 = json_data['test_results']['analysis_values']['precal']['standard_7']['img3'][0]
                buffer_red_16x_w2_p3 = json_data['test_results']['analysis_values']['precal']['standard_7']['img3'][1]
                buffer_red_16x_w3_p3 = json_data['test_results']['analysis_values']['precal']['standard_7']['img3'][2]

                buffer_red_32x_w1 = json_data['test_results']['analysis_values']['precal']['standard_8']['img1'][0]
                buffer_red_32x_w2 = json_data['test_results']['analysis_values']['precal']['standard_8']['img1'][1]
                buffer_red_32x_w3 = json_data['test_results']['analysis_values']['precal']['standard_8']['img1'][2]
                buffer_red_32x_w1_p2 = json_data['test_results']['analysis_values']['precal']['standard_8']['img2'][0]
                buffer_red_32x_w2_p2 = json_data['test_results']['analysis_values']['precal']['standard_8']['img2'][1]
                buffer_red_32x_w3_p2 = json_data['test_results']['analysis_values']['precal']['standard_8']['img2'][2]
                buffer_red_32x_w1_p3 = json_data['test_results']['analysis_values']['precal']['standard_8']['img3'][0]
                buffer_red_32x_w2_p3 = json_data['test_results']['analysis_values']['precal']['standard_8']['img3'][1]
                buffer_red_32x_w3_p3 = json_data['test_results']['analysis_values']['precal']['standard_8']['img3'][2]

                buffer_red_64x_w1 = json_data['test_results']['analysis_values']['precal']['standard_9']['img1'][0]
                buffer_red_64x_w2 = json_data['test_results']['analysis_values']['precal']['standard_9']['img1'][1]
                buffer_red_64x_w3 = json_data['test_results']['analysis_values']['precal']['standard_9']['img1'][2]
                buffer_red_64x_w1_p2 = json_data['test_results']['analysis_values']['precal']['standard_9']['img2'][0]
                buffer_red_64x_w2_p2 = json_data['test_results']['analysis_values']['precal']['standard_9']['img2'][1]
                buffer_red_64x_w3_p2 = json_data['test_results']['analysis_values']['precal']['standard_9']['img2'][2]
                buffer_red_64x_w1_p3 = json_data['test_results']['analysis_values']['precal']['standard_9']['img3'][0]
                buffer_red_64x_w2_p3 = json_data['test_results']['analysis_values']['precal']['standard_9']['img3'][1]
                buffer_red_64x_w3_p3 = json_data['test_results']['analysis_values']['precal']['standard_9']['img3'][2]

                #Buffer Green 7,8,9
                buffer_green_16x_w1 = json_data['test_results']['analysis_values']['precal']['standard_7']['img1'][3]
                buffer_green_16x_w2 = json_data['test_results']['analysis_values']['precal']['standard_7']['img1'][4]
                buffer_green_16x_w3 = json_data['test_results']['analysis_values']['precal']['standard_7']['img1'][5]
                buffer_green_16x_w1_p2 = json_data['test_results']['analysis_values']['precal']['standard_7']['img2'][3]
                buffer_green_16x_w2_p2 = json_data['test_results']['analysis_values']['precal']['standard_7']['img2'][4]
                buffer_green_16x_w3_p2 = json_data['test_results']['analysis_values']['precal']['standard_7']['img2'][5]
                buffer_green_16x_w1_p3 = json_data['test_results']['analysis_values']['precal']['standard_7']['img3'][3]
                buffer_green_16x_w2_p3 = json_data['test_results']['analysis_values']['precal']['standard_7']['img3'][4]
                buffer_green_16x_w3_p3 = json_data['test_results']['analysis_values']['precal']['standard_7']['img3'][5]

                buffer_green_32x_w1 = json_data['test_results']['analysis_values']['precal']['standard_8']['img1'][3]
                buffer_green_32x_w2 = json_data['test_results']['analysis_values']['precal']['standard_8']['img1'][4]
                buffer_green_32x_w3 = json_data['test_results']['analysis_values']['precal']['standard_8']['img1'][5]
                buffer_green_32x_w1_p2 = json_data['test_results']['analysis_values']['precal']['standard_8']['img2'][3]
                buffer_green_32x_w2_p2 = json_data['test_results']['analysis_values']['precal']['standard_8']['img2'][4]
                buffer_green_32x_w3_p2 = json_data['test_results']['analysis_values']['precal']['standard_8']['img2'][5]
                buffer_green_32x_w1_p3 = json_data['test_results']['analysis_values']['precal']['standard_8']['img3'][3]
                buffer_green_32x_w2_p3 = json_data['test_results']['analysis_values']['precal']['standard_8']['img3'][4]
                buffer_green_32x_w3_p3 = json_data['test_results']['analysis_values']['precal']['standard_8']['img3'][5]

                buffer_green_64x_w1 = json_data['test_results']['analysis_values']['precal']['standard_9']['img1'][3]
                buffer_green_64x_w2 = json_data['test_results']['analysis_values']['precal']['standard_9']['img1'][4]
                buffer_green_64x_w3 = json_data['test_results']['analysis_values']['precal']['standard_9']['img1'][5]
                buffer_green_64x_w1_p2 = json_data['test_results']['analysis_values']['precal']['standard_9']['img2'][3]
                buffer_green_64x_w2_p2 = json_data['test_results']['analysis_values']['precal']['standard_9']['img2'][4]
                buffer_green_64x_w3_p2 = json_data['test_results']['analysis_values']['precal']['standard_9']['img2'][5]
                buffer_green_64x_w1_p3 = json_data['test_results']['analysis_values']['precal']['standard_9']['img3'][3]
                buffer_green_64x_w2_p3 = json_data['test_results']['analysis_values']['precal']['standard_9']['img3'][4]
                buffer_green_64x_w3_p3 = json_data['test_results']['analysis_values']['precal']['standard_9']['img3'][5]

                #Target Green Standards 1,2,3
                target_green_std123_w1 = json_data['test_results']['ref_values']['green'][0]
                target_green_std123_w2 = json_data['test_results']['ref_values']['green'][1]
                target_green_std123_w3 = json_data['test_results']['ref_values']['green'][2]

                #Target Red Standards 4,5,6
                target_red_std123_w1 = json_data['test_results']['ref_values']['red'][0]
                target_red_std123_w2 = json_data['test_results']['ref_values']['red'][1]
                target_red_std123_w3 = json_data['test_results']['ref_values']['red'][2]

                #Normalizers red
                r_norm_w1 = json_data['test_results']['spec_values']['data']['normalizers']['r_norm'][0]
                r_norm_w2 = json_data['test_results']['spec_values']['data']['normalizers']['r_norm'][1]
                r_norm_w3 = json_data['test_results']['spec_values']['data']['normalizers']['r_norm'][2]

                #Normalizers green
                g_norm_w1 = json_data['test_results']['spec_values']['data']['normalizers']['g_norm'][0]
                g_norm_w2 = json_data['test_results']['spec_values']['data']['normalizers']['g_norm'][1]
                g_norm_w3 = json_data['test_results']['spec_values']['data']['normalizers']['g_norm'][2]

                #Red Linearity
                l_red_w1 = json_data['test_results']['spec_values']['spec_results']['well_1']['lin_red']
                l_red_w2 = json_data['test_results']['spec_values']['spec_results']['well_2']['lin_red']
                l_red_w3 = json_data['test_results']['spec_values']['spec_results']['well_3']['lin_red']

                #Green Linearity
                l_green_w1 = json_data['test_results']['spec_values']['spec_results']['well_1']['lin_green']
                l_green_w2 = json_data['test_results']['spec_values']['spec_results']['well_2']['lin_green']
                l_green_w3 = json_data['test_results']['spec_values']['spec_results']['well_3']['lin_green']

        dummy = 0 #dummy variable for printing the photo ID and Well Number
        output_df.loc[row_index] = [hostname_data, ref_unit_data, batch_code_data, date_time_data, int(dummy+1), l_red_w1, l_green_w1, r_norm_w1, g_norm_w1, well_1_center, red_16x_w1, red_32x_w1, red_64x_w1, red_16x_w1_std4, red_32x_w1_std5, red_64x_w1_std6, green_16x_w1, green_32x_w1, green_64x_w1, green_16x_w1_std4, green_32x_w1_std5, green_64x_w1_std6, buffer_red_16x_w1, buffer_red_32x_w1, buffer_red_64x_w1, buffer_green_16x_w1, buffer_green_32x_w1, buffer_green_64x_w1, target_green_std123_w1, target_red_std123_w1, int(dummy+1)]        
        output_df.loc[row_index+1] = [hostname_data, ref_unit_data, batch_code_data, date_time_data, int(dummy+2), l_red_w2, l_green_w2, r_norm_w2, g_norm_w2, well_2_center, red_16x_w2, red_32x_w2, red_64x_w2, red_16x_w2_std4, red_32x_w2_std5, red_64x_w2_std6, green_16x_w2, green_32x_w2, green_64x_w2, green_16x_w2_std4, green_32x_w2_std5, green_64x_w2_std6, buffer_red_16x_w2, buffer_red_32x_w2, buffer_red_64x_w2, buffer_green_16x_w2, buffer_green_32x_w2, buffer_green_64x_w2, target_green_std123_w2, target_red_std123_w2, int(dummy+1)]
        output_df.loc[row_index+2] = [hostname_data, ref_unit_data, batch_code_data, date_time_data, int(dummy+3), l_red_w3, l_green_w3, r_norm_w3, g_norm_w3, well_3_center, red_16x_w3, red_32x_w3, red_64x_w3, red_16x_w3_std4, red_32x_w3_std5, red_64x_w3_std6, green_16x_w3, green_32x_w3, green_64x_w3, green_16x_w3_std4, green_32x_w3_std5, green_64x_w3_std6, buffer_red_16x_w3, buffer_red_32x_w3, buffer_red_64x_w3, buffer_green_16x_w3, buffer_green_32x_w3, buffer_green_64x_w3, target_green_std123_w3, target_red_std123_w3, int(dummy+1)]

        output_df.loc[row_index+3] = [hostname_data, ref_unit_data, batch_code_data, date_time_data, int(dummy+1), l_red_w1, l_green_w1, r_norm_w1, g_norm_w1, well_1_center, red_16x_w1_p2, red_32x_w1_p2, red_64x_w1_p2, red_16x_w1_p2_std4, red_32x_w1_p2_std5, red_64x_w1_p2_std6, green_16x_w1_p2, green_32x_w1_p2, green_64x_w1_p2, green_16x_w1_p2_std4, green_32x_w1_p2_std5, green_64x_w1_p2_std6, buffer_red_16x_w1_p2, buffer_red_32x_w1_p2, buffer_red_64x_w1_p2, buffer_green_16x_w1_p2, buffer_green_32x_w1_p2, buffer_green_64x_w1_p2, target_green_std123_w1, target_red_std123_w1, int(dummy+2)]
        output_df.loc[row_index+4] = [hostname_data, ref_unit_data, batch_code_data, date_time_data, int(dummy+2), l_red_w2, l_green_w2, r_norm_w2, g_norm_w2, well_2_center, red_16x_w2_p2, red_32x_w2_p2, red_64x_w2_p2, red_16x_w2_p2_std4, red_32x_w2_p2_std5, red_64x_w2_p2_std6, green_16x_w2_p2, green_32x_w2_p2, green_64x_w2_p2, green_16x_w2_p2_std4, green_32x_w2_p2_std5, green_64x_w2_p2_std6, buffer_red_16x_w2_p2, buffer_red_32x_w2_p2, buffer_red_64x_w2_p2, buffer_green_16x_w2_p2, buffer_green_32x_w2_p2, buffer_green_64x_w2_p2, target_green_std123_w2, target_red_std123_w2, int(dummy+2)]
        output_df.loc[row_index+5] = [hostname_data, ref_unit_data, batch_code_data, date_time_data, int(dummy+3), l_red_w3, l_green_w3, r_norm_w3, g_norm_w3, well_3_center, red_16x_w3_p2, red_32x_w3_p2, red_64x_w3_p2, red_16x_w3_p2_std4, red_32x_w3_p2_std5, red_64x_w3_p2_std6, green_16x_w3_p2, green_32x_w3_p2, green_64x_w3_p2, green_16x_w3_p2_std4, green_32x_w3_p2_std5, green_64x_w3_p2_std6, buffer_red_16x_w3_p2, buffer_red_32x_w3_p2, buffer_red_64x_w3_p2, buffer_green_16x_w3_p2, buffer_green_32x_w3_p2, buffer_green_64x_w3_p2, target_green_std123_w3, target_red_std123_w3, int(dummy+2)]

        output_df.loc[row_index+6] = [hostname_data, ref_unit_data, batch_code_data, date_time_data, int(dummy+1), l_red_w1, l_green_w1, r_norm_w1, g_norm_w1, well_1_center, red_16x_w1_p3, red_32x_w1_p3, red_64x_w1_p3, red_16x_w1_p3_std4, red_32x_w1_p3_std5, red_64x_w1_p3_std6, green_16x_w1_p3, green_32x_w1_p3, green_64x_w1_p3, green_16x_w1_p3_std4, green_32x_w1_p3_std5, green_64x_w1_p3_std6, buffer_red_16x_w1_p3, buffer_red_32x_w1_p3, buffer_red_64x_w1_p3, buffer_green_16x_w1_p3, buffer_green_32x_w1_p3, buffer_green_64x_w1_p3, target_green_std123_w1, target_red_std123_w1, int(dummy+3)]
        output_df.loc[row_index+7] = [hostname_data, ref_unit_data, batch_code_data, date_time_data, int(dummy+2), l_red_w2, l_green_w2, r_norm_w2, g_norm_w2, well_2_center, red_16x_w2_p3, red_32x_w2_p3, red_64x_w2_p3, red_16x_w2_p3_std4, red_32x_w2_p3_std5, red_64x_w2_p3_std6, green_16x_w2_p3, green_32x_w2_p3, green_64x_w2_p3, green_16x_w2_p3_std4, green_32x_w2_p3_std5, green_64x_w2_p3_std6, buffer_red_16x_w2_p3, buffer_red_32x_w2_p3, buffer_red_64x_w2_p3, buffer_green_16x_w2_p3, buffer_green_32x_w2_p3, buffer_green_64x_w2_p3, target_green_std123_w2, target_red_std123_w2, int(dummy+3)]
        output_df.loc[row_index+8] = [hostname_data, ref_unit_data, batch_code_data, date_time_data, int(dummy+3), l_red_w3, l_green_w3, r_norm_w3, g_norm_w3, well_3_center, red_16x_w3_p3, red_32x_w3_p3, red_64x_w3_p3, red_16x_w3_p3_std4, red_32x_w3_p3_std5, red_64x_w3_p3_std6, green_16x_w3_p3, green_32x_w3_p3, green_64x_w3_p3, green_16x_w3_p3_std4, green_32x_w3_p3_std5, green_64x_w3_p3_std6, buffer_red_16x_w3_p3, buffer_red_32x_w3_p3, buffer_red_64x_w3_p3, buffer_green_16x_w3_p3, buffer_green_32x_w3_p3, buffer_green_64x_w3_p3, target_green_std123_w3, target_red_std123_w3, int(dummy+3)]
        

def process_folder(folder_path):
    tar_files = glob.glob(os.path.join(folder_path, "**/*CalData*.tar.gz"), recursive=True)
    output_df = pd.DataFrame(columns=["Cube-HN", "Ref-Cube-HN", "Batchcode", "Date-Time", "Well-Number", "Red-Linearity", "Green-Linearity", "Red-Norm", "Green-Norm", "Well Center Coordinatess", "Red-16X-Standard-1", "Red-32X-Standard-2", "Red-64X-Standard-3", "Red-16X-Standard-4", "Red-32X-Standard-5", "Red-64X-Standard-6", "Green-16X-Standard-1", "Green-32X-Standard-2", "Green-64X-Standard-3", "Green-16X-Standard-4", "Green-32X-Standard-5", "Green-64X-Standard-6", "Buffer-Red-Standard-7", "Buffer-Red-Standard-8", "Buffer-Red-Standard-9", "Buffer-Green-Standard-7", "Buffer-Green-Standard-8", "Buffer-Green-Standard-9", "Target(Green-Std-1,2,3)", "Target(Red-Std-4,5,6)", "Photo-ID"])
    for i, tar_file in enumerate(tar_files):
        print(tar_file)
        process_tar_file(tar_file, output_df, i*9)
    
    return output_df

if __name__ == "__main__":
    folder_path = input("Enter the folder path: ")
    df = process_folder(folder_path)
    print(df)
    output_file = "C:\\Users\\UKumar\\Desktop\\New-PreCal.xlsx"
    df.to_excel(output_file, index=False)
   