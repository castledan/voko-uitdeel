import streamlit as st
import pandas as pd
import numpy as np
import openpyxl as op
from io import BytesIO
from from_poll_to_schedule import from_poll_to_schedule

st.title("VOKO schedule preparation")

other_option_text = "Add another name..."

survey_results_uploaded = False
uitdeel_uploaded = False

st.markdown("## File upload")

upload_survey = st.file_uploader("Upload the results of the framadate survey")
if upload_survey is not None:
    try:
        survey_results = pd.read_csv(upload_survey)
        survey_results_uploaded = True
    except:
        st.error("File not valid. Check that the file is in csv format.")

upload_uitdeellist = st.file_uploader("Upload the list of uitdeel members (in csv)")
if upload_uitdeellist is not None:
    try:
        volunteer_csv = pd.read_csv(upload_uitdeellist)
        uitdeel_uploaded = True
    except:
        st.error("File not valid. Check that the file is in csv format.")

if survey_results_uploaded and uitdeel_uploaded:

    st.markdown("## Compare names from survey (left) and volunteer list (right)")

    input_list = sorted(survey_results.iloc[:, 0].dropna().tolist())

    volunteer_csv = volunteer_csv[
        volunteer_csv[volunteer_csv.columns[5]]==True
        ].reset_index(drop=True)
    coordinators_csv = volunteer_csv[
        volunteer_csv[volunteer_csv.columns[4]]==True
        ].reset_index(drop=True)
    option_list = [
        name+" "+surname for (name, surname) in zip(
            [v.strip() for v in volunteer_csv.iloc[:, 0].dropna().tolist()], 
            [v.strip() for v in volunteer_csv.iloc[:, 1].dropna().tolist()]
            )
        ]
    coordinator_list = [
        name+" "+surname for (name, surname) in zip(
            [v.strip() for v in coordinators_csv.iloc[:, 0].dropna().tolist()], 
            [v.strip() for v in coordinators_csv.iloc[:, 1].dropna().tolist()]
            )
        ] 

    selections = {}
    other_options = {}
    rows = {}
    cols0 = {}
    cols1 = {}
    cols2 = {}
    cols3 = {}
    cols4 = {}

    def check_names(selections, other_options):
        check_list = []
        for i in range(len(selections)):
            check_list.append(selections[str(i)] != "" and \
                (selections[str(i)] != other_option_text or \
                    (selections[str(i)] == other_option_text and \
                        other_options[str(i)] != "")
                )
            )
        return check_list

    def output_names_and_dict(input_list, selections, other_options):
        output_list = []
        new_name_list = []
        name_dict = {}
        for i in range(len(selections)):
            if selections[str(i)] != other_option_text:
                selected_name = selections[str(i)]
            else:
                selected_name = other_options[str(i)].strip()
                new_name_list.append(selected_name)
            output_list.append(selected_name)
            name_dict[input_list[i]] = selected_name
        return output_list, new_name_list, name_dict

    def to_excel(df, df2):
        output = BytesIO()
        book = op.load_workbook('schedule_with_macro.xlsm', keep_vba = True)
        writer = pd.ExcelWriter(output, engine='openpyxl')
        writer.book = book # Hand over input workbook
        writer.sheets.update(dict((ws.title, ws) for ws in book.worksheets))
        writer.vba_archive = book.vba_archive # Hand over VBA information 

        df.to_excel(writer,
                    header = False, index = False,
                    startrow = 0, startcol = 0)

        df2.to_excel(writer,
                    header = True, index = False,
                    startrow = 0, startcol = 8)  
        writer.save()
        processed_data = output.getvalue()
        return processed_data

    for i in range(len(input_list)):
        rows[str(i)] = st.container()
        with rows[str(i)]:
            cols0[str(i)], cols1[str(i)], cols2[str(i)], cols3[str(i)] = \
                st.columns([0.2, 2, 2, 2])
        with cols0[str(i)]:
            num_str = "&nbsp;&nbsp;&nbsp;"+str(i+1) if i+1<10 else str(i+1)
            st.markdown(
                """
                    <div style='margin-top:0.2cm;'>
                    %s.
                    </div>
                """ % num_str,
                unsafe_allow_html=True
                )
        with cols1[str(i)]:
            st.markdown(
                """
                    <div style='margin-top:0.2cm;'>
                    %s
                    </div>
                """ % input_list[i],
                unsafe_allow_html=True
                )
        with cols2[str(i)]:
            options = [""] + \
                [o for o in option_list if o not in list(selections.values())]
            if other_option_text not in options:
                options.insert(0, other_option_text)
            index_default = 1
            input_formatted = input_list[i].lower().strip()
            input_first = input_formatted.split(" ")[0] \
                if " " in input_formatted else input_formatted
            options_formatted = [o.lower().strip() for o in options]
            options_names = [o.split(" ")[0] if " " in o else o \
                for o in options_formatted]
            options_surnames = [o.split(" ")[1] if " " in o else o \
                for o in options_formatted]
            if input_formatted in options_formatted:
                index_default = options_formatted.index(input_formatted)
            if input_first in options_names:
                index_default = options_names.index(input_first)
            if input_first in options_surnames:
                index_default = options_surnames.index(input_first)
            selections[str(i)] = st.selectbox(
                label="Select name",
                index=index_default,
                options=options,
                label_visibility="collapsed", 
                key='sel_'+str(i))
        with cols3[str(i)]:
            if selections[str(i)] == other_option_text: 
                other_options[str(i)] = st.text_input(
                    label="Enter your other option...",
                    label_visibility="collapsed",
                    key='text_'+str(i)
                    )

    st.markdown(" ")
    confirm = st.button("Confirm names")

    if confirm:
        check_list = check_names(selections, other_options)
        if not all(check_list):
            st.error(
                """
                    Make sure that all names were matched correctly. There seems 
                    to be a problem with name number %s.
                """ % str(check_list.index(False)+1)
                )
        else:
            st.success('Names saved')
            output_names, new_name_list, output_dict = output_names_and_dict(
                input_list, 
                selections, 
                other_options
                )

            if new_name_list:
                st.markdown(
                    """
                        Some names that appear in the survey results were not present
                        in the volunteer list:<br><br>
                        %s
                    """ % "<br>".join(new_name_list),
                    unsafe_allow_html=True
                )

            survey_results[survey_results.columns[0]] = \
                [np.nan]+[output_dict[name] \
                    for name in survey_results[survey_results.columns[0]][1:]]
            df = from_poll_to_schedule(survey_results)

            df2_cols = ["Coordinator", "Responded", "Shift", "Name", "Assigned dates", "Comment"]

            all_names = sorted(option_list+new_name_list)
            responded = ["x" if n in output_names else "" for n in all_names]
            coordinator = ["x" if n in coordinator_list else "" for n in all_names]
            shift = [f"=CountBoldandMatch($B$3:$F$275, $L{i+2})" \
                for i in range(len(all_names))]
            assigned = [f"=FindAllDatesBold($B$3:$F$275, $L{i+2})" \
                for i in range(len(all_names))]
            comment = [""]*len(all_names)

            df2 = pd.DataFrame(
                list(zip(coordinator, responded, shift, all_names, assigned, comment)),
                columns=df2_cols
                )

            df_xlsm = to_excel(df, df2)

            st.markdown("## Download schedule excel file")

            st.download_button(
                label='📥 Download excel file',
                data=df_xlsm,
                file_name= 'schedule_output.xlsm'
                )