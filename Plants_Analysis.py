# K = ord apg iv (index 10), d = classis l (index 3), e = ordo l (index 4)
import pandas as pds
plants = pds.read_excel('species_plantarum_1753.xlsx')
dtf_size = len(plants['NOMEN'])
res_dict = {}
for i in range(0, dtf_size):
    current_pair = '{0}-{1}'.format(plants.iloc[i][3], plants.iloc[i][4])
    keys = list(res_dict.keys())
    if (keys.count(current_pair) == 0):
        temp_list = []
        temp_list.append({plants.iloc[i][10] : 1})
        res_dict[current_pair] = temp_list
    else: 
        proc_flag = 0
        for index, pair in enumerate(res_dict[current_pair]):
            key = list(pair.keys())[0]
            if (key == plants.iloc[i][10]):
                proc_flag = 1
                res_dict[current_pair][index][key] += 1
        if (proc_flag == 0):
            res_dict[current_pair].append({plants.iloc[i][10] : 1})
for pair in res_dict:
    tmp_sum = 0
    for ordo in res_dict[pair]:
        tmp_key = list(ordo.keys())[0]
        tmp_sum += ordo[tmp_key]
    for index, ordo in enumerate(res_dict[pair]):
        tmp_key = list(ordo.keys())[0]
        res_dict[pair][index][tmp_key] =format((res_dict[pair][index][tmp_key] / tmp_sum)*100, '.2f') + '%'
pairs = list(res_dict.keys())
results = []
for pair in pairs:
    tmp_string = ''
    for index, ordo in enumerate(res_dict[pair]):        
        tmp_key = list(ordo.keys())[0]
        tmp_string += '{0}: {1}, '.format(tmp_key, ordo[tmp_key])
    tmp_string = tmp_string[0:len(tmp_string) - 2]
    results.append(tmp_string)
ans_dtf = pds.DataFrame({'Pair' :  pairs, 'Results' : results})
ans_dtf.to_excel('plants_sorted.xlsx')