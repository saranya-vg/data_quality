# Importing all the necessary packages
import os
import pandas as pd
import numpy as np
from openpyxl import load_workbook
import msoffcrypto
import streamlit as st
import streamlit.components.v1 as components
from io import BytesIO
from ydata_profiling import ProfileReport
from pygwalker.api.streamlit import StreamlitRenderer
from spire.xls import *
from spire.xls.common import *


def decrypt_excel(file , password):
    """
    This function is used to decrypt the password protected file 

    parameters:
    1. file - The excel file to decrypt
    2. password - The password of the excel file

    return:
    1. decrypted - The file after decrypting
    """

    decrypted = BytesIO()
    excel_file = msoffcrypto.OfficeFile(file)
    excel_file.load_key(password=password)
    excel_file.decrypt(decrypted)

    return decrypted

def is_password_protected(file):
    """
    This function is used to check if the excel file protected or not using msoffcrypto

    parameters:
        1. file- The excel file to be checked

    returns:
        1. Boolean
    """
    try:
        office_file = msoffcrypto.OfficeFile(file)    
        return office_file.is_encrypted()
    except Exception as e:
        st.error(f'Error checking file encryption: {str(e)}')


def load_excel(file , password = None):
    """
    This function is used to load the excel file

    parameters: 
        1. file- The excel file to be loaded
        2. password- The password of the excel file

    return:
        1. df- The dataframe of the excel
    """
    try:
        protection_check = is_password_protected(file)

        if protection_check == 'invalid_format':
            return None , 'invalid_format'
        
        # code if it is a password protected file
        if protection_check:
            if not password:
                return None, 'password_protected'

            decrypted_file = decrypt_excel(file , password)
            df = pd.read_excel(decrypted_file , engine = 'openpyxl')
            return df , None
        # code if it is not a password protected file
        else:
            df = pd.read_excel(file , engine = 'openpyxl')
            return df , None

    except Exception as e:
        st.write(str(e))
        if "Workbook is password protected" in str(e):
            return None , "password_protected"
        elif 'Invalid password' in str(e):
            return None , 'invalid_password'
        elif 'Excel file format' in str(e):
            return None , 'invalid_format'
        else:
            return None , str(e)
    except Exception as e:
        return None , str(e)


def handle_duplicates(df , dataset_name):
    """
    This function is used to check and remove duplicates based on the user input

    parameters:
        1. df - The dataframe to be used
        2. dataset_name - The name of the dataset

    returns:
        1. df- The dataframe after removing duplicates
    """
    try:
        st.write(f"Checking duplicates in {dataset_name}...")
        num_duplicates = df.duplicated().sum()

        # checking for duplicated and removing them if the used selects 'Yes'
        if num_duplicates > 0:
            st.write(f"Number of duplicated rows: {num_duplicates}")
            remove_duplicates = st.radio(f"Do you want to remove duplicates from {dataset_name}?" , ('No' , 'Yes'))
            if remove_duplicates == 'Yes':
                df = df.drop_duplicates()
                st.write(f'Duplicates removed. {len(df)} rows remaining')
        else:
            st.write(f"No duplicates found!")
        return df
    
    except Exception as e:
        st.sidebar.error(f"An error occured while handling duplicates {dataset_name}: {str(e)}")
        return df

def load_dataset(suffix):
    """
    This function is used to load the dataset from the user

    parameters:
        1. suffix - It is used to make the unique key for all the widgets

    returns:
        1. df - The dataframe after removing duplicates
        2. dataset_name - The name of the dataset 
    """

    # Uploading the file and getting the name
    file = st.sidebar.file_uploader('Upload dataset' , key=f"file_uploader_{suffix}")
    dataset_name = st.sidebar.text_input('Enter a name for the dataset' , key=f"name_input_{suffix}")


    if file is not None and dataset_name != "":
        try:
            # Calling the load excel function
            df , status = load_excel(file)

            if status == 'password_protected':
                st.sidebar.warning(f'{dataset_name} is password protected. Please provide the password below.')
                password = st.sidebar.text_input(f'Enter password for {dataset_name}:' , type = 'password' , key=f"password_input_{suffix}")

                # Calling the load excel function
                if password:
                    df , status = load_excel(file , password = password)

                    if status == 'invalid_password':
                        st.sidebar.error('Invalid Password! Please try again.')
                        return None, None
                    
            if status == 'invalid_format':
                st.error(f"Invalid file format. Please upload a valid Excel file.")
                    
            if df is not None:
                st.subheader(f'Displaying first 50 records of {dataset_name}:')
                st.dataframe(df.head(50))
                # Calling the handle duplicates function
                df = handle_duplicates(df , dataset_name)
                return df,dataset_name
            
            elif password != "":
                st.sidebar.error(f"Failed to load {dataset_name}. Please check if the password is correct.")

        except Exception as e:
            st.sidebar.error(f"An error occured while loading {dataset_name}: {str(e)}")
            return None, None
        
    return None, None

def select_filters(df , key_suffix , dataset_name):
    """
    This function is used to select filter columns and values 

    parameters:
        1. df - The dataframe used to filter
        2. key_suffix - The value used to create a unique key
        3. dataset_name - The name of the dataset

    returns:
        1. filter_columns - The dictionary containing all the details to filter
    """


    filter_columns = {}
    available_columns = df.columns.to_list()
    if f"selected_columns_{key_suffix}" not in st.session_state:
        st.session_state[f"selected_columns_{key_suffix}"] = []
    
    # Selecting the columns and storing it in the session state
    selected_columns = st.multiselect(f"Select columns to filter in {dataset_name}:" , available_columns , key=f'columns_{key_suffix}' ,
                                       default=st.session_state[f"selected_columns_{key_suffix}"])
    
    st.session_state[f"selected_columns_{key_suffix}"] = selected_columns

    # For each column checking the datatype and using widgets accordingly and getting all the required fields to apply filter
    for column in selected_columns:

        column_type = df[column].dtype

        if pd.api.types.is_numeric_dtype(column_type):
            min_val , max_val = df[column].min() , df[column].max()
            # Select values to filter
            selected_range = st.slider(f"Select range for {column} in {dataset_name}:" , min_val , max_val  , (min_val , max_val) , key= f"range_{column}_{key_suffix}")
            # Getting the filter mode
            filter_mode = st.radio(f"Filter mode for {column} in {dataset_name}:" , ["Include" , "Exclude"],
                                    key= f"mode_{column}_{key_suffix}")
            filter_columns[column] = {
                'values' : selected_range,
                'type' : 'numeric',
                'mode' : filter_mode
            }
        elif pd.api.types.is_datetime64_any_dtype(column_type):
            min_date , max_date = df[column].min() , df[column].max()
            # Select values to filter
            start_date = st.date_input(f"Start date for {column} in {dataset_name}:" , min_date , key=f'start_date_{column}_{key_suffix}')
            end_date = st.date_input(f"End date for {column}:" , max_date , key=f'end_date_{column}_{key_suffix}')
            # Getting the filter mode
            filter_mode = st.radio(f"Filter mode for {column} in {dataset_name}:" , ["Include" , "Exclude"],
                                    key= f"mode_{column}_{key_suffix}")
            filter_columns[column] = {
                'values' : (start_date , end_date),
                'type' : 'date',
                'mode' : filter_mode
            }
        else:
            unique_values = list(df[column].unique())
            if f"selected_values_{column}_{key_suffix}" not in st.session_state:
                st.session_state[f"selected_values_{column}_{key_suffix}"] = []
            # Select values to filter
            selected_values = st.multiselect(f"Select values to filter in {column} in {dataset_name}:" , unique_values,
                                            key=f"values_{column}_{key_suffix}",
                                            default=st.session_state[f"selected_values_{column}_{key_suffix}"])    

            # Getting the filter mode
            filter_mode = st.radio(f"Filter mode for {column} in {dataset_name}:" , ["Include" , "Exclude"],
                                key= f"mode_{column}_{key_suffix}")
            
            filter_columns[column] = {
                'values' : selected_values,
                'mode' : filter_mode,
                'type' : 'object'
        }

    return filter_columns        


def apply_filters(df , filter_columns , dataset_name):
    """
    This function is used to apply filters selected 

    parameters:
        1. df - The dataframe used to apply the filters
        2. filter_columns - The dictionary containg all the details
        3. dataset_name - The name of the dataset

    returns:
        1. df - The dataframe after applying all the filters
    """
    try:

        # Applying the filters for each column type
        for column, filter_details in filter_columns.items():
            col_type = filter_details['type']
            filter_mode = filter_details['mode']

            if col_type == 'numeric':
                min_val , max_val = filter_details['values']
                if filter_mode == 'Include':
                    df = df[(df[column] >= min_val) & (df[column] <= max_val)]
                else:
                    df = df[~((df[column] >= min_val) & (df[column] <= max_val))]
            
            elif col_type ==  'date':
                start_date , end_date = filter_details['values']
                if filter_mode == 'Include':
                    df = df[(pd.to_datetime(df[column]).dt.date >= start_date) & (pd.to_datetime(df[column]).dt.date <= end_date)]
                else:
                    df = df[~((pd.to_datetime(df[column]).dt.date >= start_date) & (pd.to_datetime(df[column]).dt.date <= end_date))]
            
            else:
                selected_values = filter_details['values']
                filter_mode = filter_details['mode']
                if selected_values:
                    if filter_mode == 'Include':
                        df = df[df[column].isin(selected_values)]
                    elif filter_mode == 'Exclude':
                        df = df[~df[column].isin(selected_values)]

        return df    
    except Exception as e:
        st.error(f"Error while applying filters: {str(e)}")
        return df
    

def merge_data(df1_filtered , df2_filtered , dataset1_name , dataset2_name , filter_button_state):
    """
    This function is used to merge the data after filtlering

    parameters:
        1. df1_filtered - The filtered dataframe 1
        2. df1_filtered - The filtered dataframe 2
        3. dataset1_name - The name of the dataset 1
        4. dataset2_name - The name of the dataset 2

    returns:
        1. merged_df - The dataframe after merging
    """

    # Displaying the list of columns to merge on 
    st.subheader("Enter the column to Merge on")
    columns1 = df1_filtered.columns.tolist()
    selected_columns1 = st.multiselect(f"Select columns to merge from {dataset1_name}", columns1 , default=None , key = 'columns_1_merge')
    
    columns2 = df2_filtered.columns.tolist()
    selected_columns2 = st.multiselect(f"Select columns to merge from {dataset2_name}", columns2 , default=None , key = 'columns_2_merge')

    for col in columns1:
        df1_filtered[col] = df1_filtered[col].astype(str)
        df1_filtered[col] = df1_filtered[col].str.strip()

    for col in columns2:
        df2_filtered[col] = df2_filtered[col].astype(str)
        df2_filtered[col] = df2_filtered[col].str.strip()
    
    # Merging both the dataframe
    # merged_df = pd.DataFrame([])
    if filter_button_state:
        if st.button("Merge" , key = 'merger_button'):
            st.session_state["merge_button"] = not st.session_state["merge_button"]
            if selected_columns1 is not [None] and selected_columns2 is not [None] and len(selected_columns1) == len(selected_columns2) and len(selected_columns1) != 0 and len(selected_columns2) != 0 and selected_columns1 != "None" and selected_columns2 != "None":
                print(df1_filtered.shape)
                print(df2_filtered.shape)
                merged_df = pd.merge(df1_filtered , df2_filtered , left_on = selected_columns1 , right_on = selected_columns2 , how = 'outer' , suffixes = [f'_{dataset1_name}' , f'_{dataset2_name}'] , indicator = 'Exists')
                merged_df = merged_df.replace({
                    'both' : f'Present in both {dataset1_name} and {dataset2_name}',
                    'left_only' : f'Present only in {dataset1_name}',
                    'right_only' : f'Present only in {dataset2_name}'
                })
                st.write('Merged Dataset')
                st.dataframe(merged_df.head(50))

                st.subheader("Summary Table")
                summary_table = merged_df['Exists'].value_counts().reset_index()
                st.dataframe(summary_table)

    return merged_df


def convert_to_excel(df):
    """
    This function is used to convert the dataframe to exportable format

    parameters:
    1. df - The dataframe to export

    returns:
    1. processed_data - The dataframe in the exportable format
    """
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)
    processed_data = output.getvalue()
    return processed_data

def generate_profile_report(df , name):
    """
    This function is used to generate a data quality report 

    parameters:
    1. df - The dataframe to be used to generate a data quality report
    2. name - The title of the report to be displayed

    returns:
    1. The report in HTML format
    """
    profile = ProfileReport(df , title = f'{name} data quality report' , explorative = True)
    return profile.to_html()

def data_merger():
    """
    This function is used to merge 2 datasets which includes applying multiple filters and all the widgets required

    parameters:
    None

    return:
    None
    """

    # Calling the load dataset function 
    df1, dataset1_name  = load_dataset(suffix="1")
    df2, dataset2_name = load_dataset(suffix="2")

    if "filter_column_1" not in st.session_state:
        st.session_state['filter_column_1'] = {}
    if "filter_column_2" not in st.session_state:
        st.session_state['filter_column_2'] = {}

    # Calling the select filters function
    if df1 is not None and df2 is not None:
        filter_columns_1 = select_filters(df1 ,"1" , dataset1_name)
        st.session_state['filter_columns_1'] = filter_columns_1

        filter_columns_2 = select_filters(df2 ,"2" , dataset2_name)
        st.session_state['filter_columns_2'] = filter_columns_2

    if "apply_filter_button" not in st.session_state:
        st.session_state["apply_filter_button"] = False

    if "merge_button" not in st.session_state:
        st.session_state["merge_button"] = False

    if "download_button" not in st.session_state:
        st.session_state["download_button"] = False

    if df1 is not None and df2 is not None:  
        # Calling the apply filters function
        df1_filtered = pd.DataFrame()
        df2_filtered = pd.DataFrame()
        # if st.button('Apply Filter' , key = 'apply_filter_button'):
        if st.button('Apply Filter'):
            st.session_state["apply_filter_button"] = not st.session_state["apply_filter_button"]
        if st.session_state["apply_filter_button"]:
            df1_filtered = apply_filters(df1 , st.session_state["filter_columns_1"] , dataset1_name)
            df2_filtered = apply_filters(df2 , st.session_state["filter_columns_2"] , dataset2_name)

            # Calling the merge data function
        merged_df = pd.DataFrame([])
        if len(df1_filtered) != 0 and len(df2_filtered) != 0:
            # merged_df = merge_data(df1_filtered, df2_filtered , dataset1_name , dataset2_name , st.session_state["apply_filter_button"])
            # Displaying the list of columns to merge on 
            st.subheader("Enter the column to Merge on")
            columns1 = df1_filtered.columns.tolist()
            selected_columns1 = st.multiselect(f"Select columns to merge from {dataset1_name}", columns1 , default=None , key = 'columns_1_merge')
            
            columns2 = df2_filtered.columns.tolist()
            selected_columns2 = st.multiselect(f"Select columns to merge from {dataset2_name}", columns2 , default=None , key = 'columns_2_merge')
            
            # Merging both the dataframe
            merged_df = pd.DataFrame([])
            if st.session_state["apply_filter_button"]:
                if st.button("Merge"):
                    st.session_state["merge_button"] = not st.session_state["merge_button"]
                if st.session_state["merge_button"]:
                    if selected_columns1 is not [None] and selected_columns2 is not [None] and len(selected_columns1) == len(selected_columns2) and len(selected_columns1) != 0 and len(selected_columns2) != 0 and selected_columns1 != "None" and selected_columns2 != "None":
                        print(df1_filtered.shape)
                        print(df2_filtered.shape)
                        merged_df = pd.merge(df1_filtered , df2_filtered , left_on = selected_columns1 , right_on = selected_columns2 , how = 'outer' , suffixes = [f'_{dataset1_name}' , f'_{dataset2_name}'] , indicator = 'Exists')
                        merged_df = merged_df.replace({
                            'both' : f'Present in both {dataset1_name} and {dataset2_name}',
                            'left_only' : f'Present only in {dataset1_name}',
                            'right_only' : f'Present only in {dataset2_name}'
                        })
                        st.write('Merged Dataset')
                        st.dataframe(merged_df[merged_df['Exists'] == f'Present in both {dataset1_name} and {dataset2_name}'].head(50))

                        st.subheader("Summary Table")
                        summary_table = merged_df['Exists'].value_counts().reset_index()
                        st.dataframe(summary_table)

        # Calling the convert to excel function
        excel_df = convert_to_excel(merged_df)

        #Getting the name of the output file and downloading the output file  
        if len(merged_df) != 0:
            output_name = st.text_input('Enter the output file name')
            if output_name != "":
                if st.session_state["apply_filter_button"] and st.session_state["merge_button"]: 
                    if st.download_button(label = "Download the merged dataset" , 
                                        data = excel_df,
                                        file_name = output_name+".xlsx" ,
                                        key='donwload_button_merger'):
                        st.session_state["download_button"] = not st.session_state["download_button"]


def classify_variable(data_type):
    """
    This function is used to classify the variable based on the type

    Parameters:
        1. data_type - The data type of the column

    Returns:
        1. str - The category of the variable
    """

    # Classifying the columns based on its data type into Numerical , Categorical , Data and Other
    if np.issubdtype(data_type , np.number):
        return "Numerical"
    elif np.issubdtype(data_type , np.datetime64):
        return "Date"
    elif data_type == "object":
        return "Categorical"
    else:
        return "Other"
    
def detect_outliers(df , column , method , threshold = 1.5):
    """
    This function is used to detect the outliers in numerical columns using IQR and Z-Score method

    Parameters:
        1. column - The name of the column to analyze
        2. method - The method to use
        3. Threshold - The value used in both the methods to find the outliers by using this as a threshold

    Returns:
        1. "Total Outliers" : The total number of outliers
        2. "Outliers (%)" : The percentage of outliers
        3. "Ouliers" : The list of outliers
    """

    # Dropping null values
    data = df[column].dropna()

    # Detecting the outliers for both IQR and Z-Score
    if len(data) != 0:
        if method == 'iqr':
            # Calculatinig the lower and upper bound
            Q1 = data.quantile(0.25)
            Q3 = data.quantile(0.75)
            IQR = Q3 - Q1
            lower_bound = Q1 - threshold * IQR
            upper_bound = Q3 + threshold * IQR
            # Extracting the outliers
            outliers = data[(data < lower_bound) | (data > upper_bound)]
        elif method == 'zscore':
            # Calculating the Z-Score
            mean = data.mean()
            std_dev = data.std()
            z_scores = (data - mean)/std_dev
            z_score_threshold = 2
            # Extracting the outliers after calculating the Z-Score and threshold
            outliers = data[(z_scores < -z_score_threshold) | (z_scores > z_score_threshold)]

        # Returning the required details
        return {
            "Total Outliers" : len(outliers),
            "Outliers (%)" : str(round((len(outliers)/len(data))*100 , 2)) + "%",
            "Ouliers" : outliers.tolist()
        }
    
    else:
        return {
            "Total Outliers" : 0,
            "Outliers (%)" : "0%",
            "Ouliers" : []
        }

def detect_data_type_inconsistencies(df , column):
    """
    This function is used to detect if there are multiple data types within the same column.

    Parameters:
        1. column - The column to analyze

    Returns:
        1. "Consistent" : Boolean to determine if it is consistent
        2. "Inconsistent Types" : All the data types present if multiple are present
        3. "Inconsistent Count" : Count of data types present
        4. "Data Types" : List of all the data types
    """

    # Extracting the unique data types present
    types = df[column].dropna().apply(type).value_counts()
    inconsistent_types = types[types > 0].index.tolist()

    if len(inconsistent_types) > 1:
        # Extracting the values of the inconsistent data types
        inconsistent_values = df[column][~df[column].apply(lambda x: isinstance(x , type(df[column].dropna().iloc[0])))]
        # Returning the required details
        return {
            "Consistent" : False,
            "Inconsistent Types" : [t.__name__ for t in types.index],
            "Inconsistent Count" : len(inconsistent_values),
            "Data Types" : inconsistent_values.dropna().unique().tolist()
        }
    
    # Returning the required details
    return {
            "Consistent" : True,
            "Inconsistent Types" : None,
            "Inconsistent Count" : 0,
            "Data Types" : None
        }


def calculate_entropy(df , col):
        """
        This function is used to calculate the entropy of each column based on the number of unique values

        Parameters:
            1. col - The column to calulate

        Returns:
            1. entropy - The entropy value of that column
        """

        value_counts = df[col].value_counts(normalize = True)
        entropy = -np.sum(value_counts*np.log2(value_counts))
        return entropy


def generate_custom_report(df , output_path = os.getcwd()):
    """
    This function is used to generate a custom data quality report

    Parameters:
        None

    Returns:
        None - It saves the report automatically in the form of an excel file
    """

    # Generating the data frame with columns and values for the overall statistics
    overall_report_columns = ["Overall Statistic" , "Value"]
    overall_statistic_list = [
        "Total Number of Columns",
        "Total Number of Rows",
        "Missing Cells",
        "Missing Cells (%)",
        "Duplicate Rows",
        "Duplicate Rows (%)"
    ]
    value_list = [
        df.shape[1],
        df.shape[0],
        df.isnull().sum().sum(),
        str(round(df.isnull().sum().sum()*100/(df.shape[0]*df.shape[1]) , 2)) + "%",
        df.duplicated().sum(),
        str(round(df.duplicated().sum()*100/df.shape[0] , 2)) + "%"
    ]
    overall_report = pd.DataFrame(columns = overall_report_columns)
    overall_report["Overall Statistic"] = overall_statistic_list 
    overall_report["Value"] = value_list 
    
    # Creating a dataframe with all the required columns and values for the column wise report
    column_wise_report = pd.DataFrame({
        "Column Name" : df.columns,
        "Data Type" : df.dtypes,
        "Variable Category" : [classify_variable(data_type) for data_type in df.dtypes],
        "Comments" : "",
        "Consistent Data Types" : [detect_data_type_inconsistencies(df , col)['Consistent'] for col in df.columns],
        "Inconsistent Data Types" : [detect_data_type_inconsistencies(df , col)['Inconsistent Types'] for col in df.columns],
        "Inconsistent Count" : [detect_data_type_inconsistencies(df , col)['Inconsistent Count'] for col in df.columns],
        "Unique Values" : [df[col].nunique() for col in df.columns],
        "Unique Values (%)" : [round(df[col].nunique()/len(df)*100 , 2) for col in df.columns], 
        "Entropy" : [round(calculate_entropy(df , col) , 2) if not np.issubdtype(df[col].dtype , np.number) else None for col in df.columns],
        "Missing Values" : [df[col].isna().sum() for col in df.columns], 
        "Missing Values (%)" : [round(df[col].isna().sum()/len(df)*100 , 2) for col in df.columns],
        "Non-Missing Values" : [df[col].notna().sum() for col in df.columns],
        "Imbalance" : [True if df[col].value_counts(normalize = True).max()*100 > 70 and not np.issubdtype(df[col].dtype , np.number) else None for col in df.columns],
        "Mean" : [df[col].mean() if np.issubdtype(df[col].dtype , np.number) else None for col in df.columns],
        "Min" : [df[col].min() if np.issubdtype(df[col].dtype , np.number) else None for col in df.columns],
        "Max" : [df[col].max() if np.issubdtype(df[col].dtype , np.number) else None for col in df.columns],
        "Range" : [df[col].max() - df[col].min() if np.issubdtype(df[col].dtype , np.number) else f"{df[col].min()} to {df[col].max()}" if np.issubdtype(df[col].dtype , np.datetime64) else None for col in df.columns],
        "Variance" : [df[col].var() if np.issubdtype(df[col].dtype , np.number) else None for col in df.columns],
        "Standared Deviation" : [df[col].std() if np.issubdtype(df[col].dtype , np.number) else None for col in df.columns],
        "Skewness" : [df[col].skew() if np.issubdtype(df[col].dtype , np.number) else None for col in df.columns],
        "Kurtosis" : [df[col].kurt() if np.issubdtype(df[col].dtype , np.number) else None for col in df.columns],
        "Outliers (IQR)" : [detect_outliers(df , col , "iqr" , 1.5)["Total Outliers"] if np.issubdtype(df[col].dtype , np.number) else None for col in df.columns],
        "Outliers (%) (IQR)" : [detect_outliers(df , col , "iqr" , 1.5)["Outliers (%)"] if np.issubdtype(df[col].dtype , np.number) else None for col in df.columns],
        "Outliers (Z-Score)" : [detect_outliers(df , col , "zscore" , 1.5)["Total Outliers"] if np.issubdtype(df[col].dtype , np.number) else None for col in df.columns],
        "Outliers (%) (Z-Score)" : [detect_outliers(df , col , "zscore" , 1.5)["Outliers (%)"] if np.issubdtype(df[col].dtype , np.number) else None for col in df.columns],
        "First Value" : [df[col].dropna().iloc[0] if df[col].notna().sum() > 0 else None for col in df.columns],
        "Last Value" : [df[col].dropna().iloc[-1] if df[col].notna().sum() > 0 else None for col in df.columns],
        "Most Frequent Value" : [df[col].mode().iloc[0] if not df[col].mode().empty else None for col in df.columns],
        "Constant Column" : [df[col].nunique() == 1 for col in df.columns]
    })

    # Adding the percentage symbol to the required columns
    column_wise_report["Unique Values (%)"] = column_wise_report["Unique Values (%)"].astype(str) + "%"
    column_wise_report["Missing Values (%)"] = column_wise_report["Missing Values (%)"].astype(str) + "%"
    column_wise_report.fillna('-' , inplace = True)

    # Resetting the index
    column_wise_report.reset_index(drop = True , inplace = True)

    # Generating a correlation matrix on the numerical columns
    numerical_columns = df.select_dtypes(include = [np.number]).columns
    correlation_matrix = df[numerical_columns].corr()


    # Looping through each row to summarize the comments
    for i in range(column_wise_report.shape[0]):
        comments = []
        # Checking for consistent columns
        if not column_wise_report.loc[i , 'Consistent Data Types']:
            # comments.append(f"Contains multiple data types : {str(column_wise_report.loc[i , 'Inconsistent Data Types'])[1:-1]}")
            comments.append(f"Inconsistent data types : {str(column_wise_report.loc[i , 'Inconsistent Data Types'])[1:-1]}")

        # Checking for Missing values
        if column_wise_report.loc[i , 'Missing Values'] != 0:
            comments.append(f"Contains {column_wise_report.loc[i , 'Missing Values']} or {column_wise_report.loc[i , 'Missing Values (%)']} missing values")
        
        # # Check for high cardinality
        # if column_wise_report.loc[i , 'Unique Values (%)'] > 25:
        #     # comments.append(f"Contains more distinct values(High Cardinality) - {column_wise_report.loc[i , 'Unique Values']} or {column_wise_report.loc[i , 'Unique Values (%)']}%")
        #     comments.append(f"Contains more distinct values - {column_wise_report.loc[i , 'Unique Values']} or {column_wise_report.loc[i , 'Unique Values (%)']}%")


        # Check for high cardinality
        CARDINALITY_THRESHOLD = round(np.log2(len(df)) , 2)
        if column_wise_report.loc[i , 'Variable Category'] == 'Object' and column_wise_report.loc[i , 'Entropy'] > CARDINALITY_THRESHOLD:
            comments.append(f"Presence of more distinct values - {column_wise_report.loc[i , 'Unique Values']} or {column_wise_report.loc[i , 'Unique Values (%)']}%")


        # Checking for constant columns
        if column_wise_report.loc[i , 'Constant Column']:
            comments.append(f"Column contains only constant values")


        # Checking for imbalance
        if column_wise_report.loc[i , 'Imbalance'] == True:
            most_frequent_percentage = round(df[column_wise_report.loc[i , 'Column Name']].value_counts(normalize = True).max()*100 , 2)
            most_frequent_value = column_wise_report.loc[i , 'Most Frequent Value']
            comments.append(f"Column in highly imbalanced with value {most_frequent_value} of {most_frequent_percentage}%")



        # Checking for Variability in numerical columns
        if column_wise_report.loc[i , 'Variable Category'] == 'Numerical' and column_wise_report.loc[i , 'Range'] == 0:
            comments.append('Numerical column has no variability')

        # # Checking for skewness in the dataset
        # if column_wise_report.loc[i , 'Variable Category'] == 'Numerical' and abs(column_wise_report.loc[i , 'Skewness']) >= 1:
        #     if column_wise_report.loc[i , 'Skewness'] >= 1:
        #         # comments.append(f"Column is highly positively skewed")
        #         comments.append(f"Outliers are present to the right of the distribution")
        #     elif column_wise_report.loc[i , 'Skewness'] <= -1:
        #         # comments.append(f"Column is highly negatively skewed")
        #         comments.append(f"Outliers are present to the right of the distribution")

        # Checking for outliers
        if column_wise_report.loc[i , 'Variable Category'] == 'Numerical' and column_wise_report.loc[i , 'Outliers (IQR)'] != 0:
            if column_wise_report.loc[i , 'Outliers (Z-Score)'] != 0:
                comments.append(f"There are {column_wise_report.loc[i , 'Outliers (IQR)']} or {column_wise_report.loc[i , 'Outliers (%) (IQR)']} outliers found using IQR method and {column_wise_report.loc[i , 'Outliers (Z-Score)']} or {column_wise_report.loc[i , 'Outliers (%) (Z-Score)']} outliers found using Z-Score method")
            else:
                comments.append(f"There are {column_wise_report.loc[i , 'Outliers (IQR)']} or {column_wise_report.loc[i , 'Outliers (%) (IQR)']} outliers found using IQR method")
        elif column_wise_report.loc[i , 'Variable Category'] == 'Numerical' and column_wise_report.loc[i , 'Outliers (Z-Score)'] != 0:
            comments.append(f"There are {column_wise_report.loc[i , 'Outliers (Z-Score)']} or {column_wise_report.loc[i , 'Outliers (%) (Z-Score)']} outliers found using IQR method")


        # Checking high correlation
        correlation_threshold = 0.8
        if column_wise_report.loc[i , 'Variable Category'] == 'Numerical':
            col_name = column_wise_report.loc[i , 'Column Name']
            correlated_cols = correlation_matrix[col_name][(correlation_matrix[col_name].abs() > correlation_threshold) & (correlation_matrix[col_name].abs() < 1)]
            if not correlated_cols.empty:
                comments.append(f"Highly correlated with: {', '.join(correlated_cols.index)}")



        # Concating the comments into a string
        if len(comments) != 0:
            summary_comment = ". ".join(comments)
            column_wise_report.loc[i , 'Comments'] = summary_comment
        else:
            column_wise_report.loc[i , 'Comments'] = "No discrepancies detected"

    return overall_report , column_wise_report , correlation_matrix


def generate_data_quality_report():
    """
    This function is used to generate a data quality report

    parameters:
    None

    returns:
    None
    """

    df, dataset_name  = load_dataset(suffix="3")
    report_file_name = ""

    if 'generate_report' not  in st.session_state:
        st.session_state['generate_report'] = False 

    if 'report_download' not in st.session_state:
        st.session_state['report_download'] = False 

    if df is not None:
        # Selecting the columns to use to generate a report
        columns = df.columns.to_list()
        columns.insert(0 , "Select All")
        selected_columns = st.multiselect(f"Select the columns to generate report" , columns , default=None)
        if "Select All" not in selected_columns:
            df = df[selected_columns]


        profile_html =""
        report_file_name = ""
        if st.button('Generate Report'):
            st.session_state['generate_report'] = not st.session_state['generate_report']
        if st.session_state['generate_report']:
            st.subheader("Generating data quality report")
            st.info("This might take a few seconds to a few minutes.")
            with st.spinner('Loading, please wait...'):
                profile_html = generate_profile_report(df , dataset_name)
                overall_report , column_wise_report , correlation_matrix = generate_custom_report(df)

            st.success('Data Quality Report Generated!')
            st.dataframe(overall_report)
            st.dataframe(column_wise_report)
            st.dataframe(correlation_matrix)
            components.html(profile_html ,height = 1000 , scrolling = True)


        # Getting the name of the output file and downloading the output file  
        
        if st.session_state['generate_report']:
            report_file_name = st.text_input('Enter the report file name' , key = "report_output_name")
            if report_file_name != "":
                with pd.ExcelWriter(os.path.join(os.getcwd() , f"{report_file_name}.xlsx") , engine = 'xlsxwriter') as writer:
                    overall_report.to_excel(writer , sheet_name = 'Overall Statistics' , index = False)
                    column_wise_report.to_excel(writer , sheet_name = 'Column Wise Statistics' , index = False)
                    correlation_matrix.to_excel(writer , sheet_name = 'Correlation Matrix')
                if st.download_button(label = "Download data quality report as HTML" , 
                                    data = profile_html,
                                    file_name = report_file_name+'.html' , 
                                    key = 'download_button_data_quality_report',
                                    mime='text/html'):
                    st.session_state['report_download'] = not st.session_state['report_download']
                                    


def data_visualization():
    """
    This function is used to visualize the data using PyGWalker

    parameters:
        None

    returns:
        None
    """

    # Loading the dataset
    df , dataset_name = load_dataset(suffix="4")
    if df is not None:
        # Rendering the interface
        pyg_app = StreamlitRenderer(df)
        pyg_app.explorer()


def load_dataset_pivot(suffix = 0):
    """
    This function is used to load the dataset from the user

    parameters:
        1. suffix - It is used to make the unique key for all the widgets

    returns:
        1. df - The dataframe after removing duplicates
        2. dataset_name - The name of the dataset 
    """
    

    # Uploading the file and getting the name
    file = st.sidebar.file_uploader('Upload dataset' , key=f"file_uploader_{suffix}")
    dataset_name = st.sidebar.text_input('Enter a name for the dataset' , key=f"name_input_{suffix}")


    if file is not None and dataset_name != "":
        
        try:
            # Creating a temporary file to load using spire
            temp_file_path = os.path.join('temp_file.xlsx')

            with open(temp_file_path , 'wb') as temp_file:
                temp_file.write(file.getbuffer())

            # Using the temp file to read using spire
            df = pd.read_excel(file)
            workbook = Workbook()
            workbook.LoadFromFile(temp_file_path)

            # Removing the temp file
            # os.remove(temp_file_path)

            return df , workbook , None

        except Exception as e:
            st.sidebar.error(f"An error occured while loading {dataset_name}: {str(e)}")
            return None, None, None
        
    return None, None, None


def select_columns_and_target(df):
    """
    This function is used to select the columns to analyse and select the target column

    parameters:
        1. df - The datframe to select the columns and target

    returns:
        1. cols_to_analyse - The list of columns to analyse
        2. target - The target column
    """

    try :

        columns = df.columns.tolist()
        cols_to_analyse = st.multiselect(f"Select columns to anlayse", columns , default=None , key = 'columns_1_analyse')
        columns.insert(0 , 'No Target')
        targets = st.multiselect(f"Select the target column", columns , default=None , key = 'columns_2_analyse')
        
        return cols_to_analyse , list(targets)
    
    except Exception as e:
        st.sidebar.error(f"An error occured while selecting columns and targets : {str(e)}")
        return None, None


def column_letter(col_idx , buffer):
    """
    This function is used to generate a letter from the given number and buffer

    parameters:
        1. col_idx - The number which will denote the number of columns
        2. buffer - The number which will be used as buffer to add some columns in between the pivot tables

    returns:
        1. letter - The string alphabet which will used in cell ranges
    """

    try : 

        letter = ""
        while col_idx > 0:
            col_idx , remainder = divmod(col_idx - 1 + buffer, 26)
            letter = chr(65 + remainder) + letter
        return letter
    
    except Exception as e:
        st.sidebar.error(f"An error occured while generating letters : {str(e)}")
        return None


def get_excel_range(df):
    """
    This function is used to get the used range using the shape of the dataframe

    parameters:
        1. df -  The dataframe to get the shape

    retuns:
        1. excel_range - The range of data used in the form of a string  
    """

    try:

        # get the shape
        rows , cols = df.shape  
        # Identify the last column
        last_column = column_letter(cols , 0)   
        excel_range = f'A1:{last_column}{rows}'

        return excel_range

    except Exception as e:
        st.sidebar.error(f"An error occured while generating excel range : {str(e)}")
        return None

def create_pivot_table_spire(piVotCache , current_sheet , col , target ,  type , start_range):
    """
    This function is used to create pivot tables based on the parameters

    parameters:
        1. piVotCache - The pivot cache which is used to create pivot tables
        2. current_sheet - The current working worksheet
        3. col - The column which will be used as the row field in the pivot table
        4. target - The column which will be used as the target or as the column field in the pivot table
        5. type - The type of the pivot table to display the values accordingly
        6. start_range - The cell where the pivot table is to be added
    """

    try :

        # Add a PivotTable to the worksheet and set the location and cache of it
        pivotTable = current_sheet.PivotTables.Add("Pivot Table", current_sheet.Range[start_range], piVotCache)

        # Set row labels
        pivotTable.Options.RowHeaderCaption = "Row Labels"

        # Drag Col field to rows area
        rowField = pivotTable.PivotFields[col]
        rowField.Axis = AxisTypes.Row

        columnField = pivotTable.PivotFields[target]
        columnField.Axis = AxisTypes.Column

        row = pivotTable.DataFields.Add(pivotTable.PivotFields[target] , f"Count of {col}", SubtotalTypes.Count)

        if type == 'row':
            row.ShowDataAs = PivotFieldFormatType.PercentageOfRow  
        elif type == 'col':
            row.ShowDataAs = PivotFieldFormatType.PercentageOfColumn     

        # Apply a built-in style to the pivot table
        pivotTable.BuiltInStyle = PivotBuiltInStyles.PivotStyleMedium10

        # Calculate data
        pivotTable.CalculateData()      

        # Soritng the pivot table
        pivotTable.SortType = PivotFieldSortType.Descending
        # row.SortType = PivotFieldSortType.Descending

        # Show subtotals
        pivotTable.ShowSubtotals = True 

        # Refresh pivot table
        pivotTable.Cache.IsRefreshOnLoad = True

        # Re-calculate data
        pivotTable.CalculateData()

    except Exception as e:
        st.sidebar.error(f"An error occured while creating pivot table : {str(e)}")
        return None


def generate_pivot_tables_spire(workbook , cols_to_analyse , targets , df):
    """
    This function is used to create a new sheet and create pivot tables into that sheet

    parameters:
        1. workbook - The workbook object that is being used
        2. cols_to_analyse - The list of columns to analyse
        3. target - The target columns
        4. df - The datframe of the used range
    """

    try: 
        # Get the first worksheet
        sheet = workbook.Worksheets[0]

        # Select the data source range dynamically
        range = get_excel_range(df)
        cellRange = sheet.Range[range]
        # st.write(range)
        piVotCache = workbook.PivotCaches.Add(cellRange)

        with st.spinner('Loading, please wait...'):
            for i , target in enumerate(targets):
                for col in cols_to_analyse:

                    # Creating a new sheet
                    current_sheet = workbook.Worksheets.Add(f"{col}_{i+1}")

                    # st.write(df[target].value_counts())

                    # Creating 3 letters to place the pivot tables
                    target_unique = df[target].nunique()
                    letter1 = 'A'
                    letter2 = column_letter(target_unique , 5)
                    letter3 = column_letter(2*target_unique , 9)

                    # Calling the function to create the pivot tables
                    create_pivot_table_spire(piVotCache , current_sheet , col , target , 'raw' , f'{letter1}1')
                    create_pivot_table_spire(piVotCache , current_sheet , col , target , 'row' , f'{letter2}1')
                    create_pivot_table_spire(piVotCache , current_sheet , col , target , 'col' , f'{letter3}1')
                    

        st.success("Pivot tables generated successfully!")

        # Saving the output with the pivot tables
        output_name = st.text_input('Enter the output file name')
        if output_name != "":
            workbook.SaveToFile(f'{output_name}.xlsx' , ExcelVersion.Version2016)
            st.success("File Saved successfully!")

    except Exception as e:
        st.sidebar.error(f"An error occured while generating pivot tables : {str(e)}")
        return None

def pivot_table_generator():
    """
    The main function for pivot table generator
    """
    
    # Load the datset
    df , workbook , wb = load_dataset_pivot()

    

    if df is not None:

        # Select the columns to anlayze and target columns
        cols_to_analyse , targets = select_columns_and_target(df)
        # st.write(targets)

        if 'generate_pivot' not  in st.session_state:
            st.session_state['generate_pivot'] = False 

        if 'pivot_download' not in st.session_state:
            st.session_state['pivot_download'] = False 

        if st.button("Generate Pivot Tables"):
            st.session_state['generate_pivot'] = not st.session_state['generate_pivot']
        if st.session_state['generate_pivot']:
            if targets[0] != "No Target":
                generate_pivot_tables_spire(workbook , cols_to_analyse , targets , df)
            else:
                pass

def main():

    # Selecting the utility and calling the respective functions
    option = st.sidebar.selectbox("Choose the Utility" , ['Data Merger' , 'Data Quality Report Generator' , 'Pivot Table Generator' , 'Data Visualization'])
    if option == 'Data Merger':
        st.title("Data Merger")
        data_merger()
    elif option == 'Data Quality Report Generator':
        st.title("Data Quality Generator")
        generate_data_quality_report()
    elif option == 'Pivot Table Generator':
        st.title("Pivot Table Generator")
        pivot_table_generator()
    else:
        st.title("Data Visualization")
        data_visualization()

    

if __name__ == "__main__":
    main()
