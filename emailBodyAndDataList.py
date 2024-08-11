def bodyCreate(data,template):
    try:
        # Format the template using the data dictionary
        masg = template.format(**data)
        return masg
    except KeyError as e:
        return f"Missing key: {e} in the data dictionary."
    

def getDataList(masg,df):
    dataList =[]
    for index, row in df.iterrows():
        # values from columns
        email = row['Email']
        dataList.append( {'email': email,'body': bodyCreate(row,masg)} )
    
    return dataList