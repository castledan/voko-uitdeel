import pandas as pd
import streamlit as st

empty_lines = 4

@st.cache_data
def from_poll_to_schedule(survey_results):
    dates = [c for c in survey_results.columns.tolist() if 'Unnamed' not in c and '.' not in c]
    hours_common = sorted(list(set([a.split('(')[0].strip() for a in survey_results.iloc[0].tolist() if pd.notnull(a)])))

    final_names = {}
    max_lens = []
    for h_c in hours_common:
        final_names[h_c] = []

    locations = []
    loc_dict = {'A': 'Averechts', 'K': 'Krachtstation'}

    initial_row = 4
    alphabeth = ['C', 'D', 'E', 'F', 'G', 'H']
    for d in dates:
        ans_date = survey_results[[survey_results.columns.tolist()[0]]+[c for c in survey_results.columns.tolist() if d in c]]
        ans_date.iat[0,0] = 'Name'
        ans_date.columns = ans_date.iloc[0]
        ans_date = ans_date[1:]
        hours = ans_date.columns.tolist()[1:]
        names = {}
        for h in hours:
            h_c = h.split('(')[0].strip()
            names[h_c] = sorted(ans_date[(ans_date[h]!='No') & (ans_date[h]!='Nee')]['Name'].tolist())
        try:
            locations.append(loc_dict[h.split('(')[1][:-1]])
        except:
            locations.append('')
        max_len = max([len(l) for l in list(names.values())])
        max_lens.append(max_len)
        for h_c in hours_common:
            names[h_c] = names[h_c]+['']*(max_len-len(names[h_c]))
        rows_to_add = names[hours_common[0]]+['']*empty_lines
        final_row = initial_row+len(rows_to_add)-1
        for h_num, h_c in enumerate(hours_common):
            let = alphabeth[h_num]
            header = f"=CountBold({let}{initial_row}:{let}{final_row})"
            final_names[h_c] += [header]+names[h_c]+['']*empty_lines
        initial_row = initial_row+len(rows_to_add)+1

    dates_col = ['']*2
    for i in range(len(dates)):
        dates_col += [dates[i],locations[i]] + ['']*(max_lens[i]+empty_lines-1)

    coord_col = ['Coordinator', '16:30 - 18:30']
    for i in range(len(dates)):
        coord_col += ['.']+['']*(max_lens[i]+empty_lines)

    hours_0_col = ['1e vroege shift', hours_common[0]] + final_names[hours_common[0]]
    hours_1_col = ['2e vroege shift', hours_common[1]] + final_names[hours_common[1]]
    hours_2_col = ['1e late shift', hours_common[2]] + final_names[hours_common[2]]
    hours_3_col = ['2e late shift', hours_common[3]] + final_names[hours_common[3]]

    df = pd.DataFrame(list(zip(dates_col, coord_col, hours_0_col, hours_1_col, hours_2_col, hours_3_col)),)

    return df