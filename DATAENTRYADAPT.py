import PySimpleGUI as sg
import pandas as pd

sg.theme('SandyBeach')

#Path of the Excel file linked to python
EXCEL_FILE=r"D:\Registration.xlsx"
df=pd.read_excel(EXCEL_FILE)

#The language column
langCol=[
[sg.Text('Can Speak',size=(15,1),key='Can Speak'),
sg.Checkbox('Hindi',key='Hindi'),
sg.Checkbox('English',key='English'),
sg.Checkbox('Gujarati',key='Gujarati'),
sg.Checkbox('Marathi',key='Marathi')],
[sg.Text('Can Read/Write',size=(15,1),key='Can Read/Write'),
sg.Checkbox('Hindi',key='Hindi'),
sg.Checkbox('English',key='English'),
sg.Checkbox('Gujarati',key='Gujarati'),
sg.Checkbox('Marathi',key='Marathi')]]

#The additonal info column
moreFrame=[[sg.Text('School',size=(15,1)),sg.InputText(key='School')],
[sg.Text('Standard',size=(15,1)),sg.InputText(key='Standard')],
[sg.Text('Academic Year',size=(15,1)),sg.InputText(key='Academic Year')],
[sg.Text('Status',size=(15,1)),sg.InputText(key='Status')],
[sg.Text('Medium',size=(15,1)),sg.InputText(key='Medium')],
[sg.Text('Interests',size=(15,1)),sg.InputText(key='Interests')],
[sg.Text('Disability',size=(15,1)),sg.InputText(key='Disability')],
[sg.Text('Category',size=(15,1)),sg.InputText(key='Category')]]


#main column
layout=[

[sg.Text('Please fill')],
[sg.Text('Name',size=(15,1)),sg.InputText(key='Name')],
[sg.Text('Age',size=(15,1)),sg.InputText(key='Age')],
[sg.In(key='DOB', enable_events=True, visible=False),
           sg.Col([[sg.CalendarButton('DOB', target='DOB', pad=None, 
                                key='-CAL1-',format=('%Y-%m-%d'))]])],
[sg.Text('Address'),sg.ML(size=(30,3),key='Address')],
[sg.Text('UDI',size=(15,1)),sg.InputText(key='UDI')],
[sg.Text('Father name',size=(15,1)),sg.InputText(key='Father name')],
[sg.Text('Father contact',size=(15,1)),sg.InputText(key='Father contact')],
[sg.Text('Father profession',size=(15,1)),sg.Combo(['Self-Employeed','Govt-Employee','Private-Employee','Other'],key='Father profession')],
[sg.Text('Father aadhaar no',size=(15,1)),sg.InputText(key='Father aadhaar no')],
[sg.Text('Father income',size=(15,1)),sg.Combo(['Rs. 0 to Rs. 2.5L','Rs. 2.5L-Rs. 5L','Rs. 5L-Rs. 7L','Rs. 7L+'],key='Father income')],
[sg.Text('Mother name',size=(15,1)),sg.InputText(key='Mother name')],
[sg.Text('Mother contact',size=(15,1)),sg.InputText(key='Mother contact')],
[sg.Text('Mother profession',size=(15,1)),sg.Combo(['Self-Employeed','Govt-Employee','Private-Employee','Other'],key='Mother profession')],
[sg.Text('Mother aadhaar no',size=(15,1)),sg.InputText(key='Mother aadhaar no')],
[sg.Text('Mother income',size=(15,1)),sg.Combo(['Rs. 0 to Rs. 2.5L','Rs. 2.5L-Rs. 5L','Rs. 5L-Rs. 7L','Rs. 7L+'],key='Mother income')],
[sg.Frame('Languages',langCol),sg.Frame('More details',moreFrame)],


[sg.Submit(),sg.Exit(),sg.Button('Clear')]

]
#resizable enabled window to be full screen
window=sg.Window('Registration',layout,resizable=True)

#to clear the inputs in all the fields
def clear_inputs():
    for key in values:
        window['Name'].update('')
        window['DOB'].update('')
        window['Age'].update('')
        window['Address'].update('')
        window['UDI'].update('')
        window['Father name'].update('')
        window['Father contact'].update('')
        window['Father profession'].update('')
        window['Father aadhaar no'].update('')
        window['Father income'].update('')
        window['Mother name'].update('')
        window['Mother contact'].update('')
        window['Mother profession'].update('')
        window['Mother aadhaar no'].update('')
        window['Mother income'].update('')
        window['Can Speak'].update('')
        window['Can Read/Write'].update('')
        window['School'].update('')
        window['Standard'].update('')
        window['Academic Year'].update('')
        window['Status'].update('')
        window['Medium'].update('')
        window['Interests'].update('')
        window['Disability'].update('')
        window['Category'].update('')
        return None



while True:
    event,values=window.read()
    if event == sg.WIN_CLOSED or event=='Exit':
        break
    if event=='Clear':
        clear_inputs()

    if event=='Submit':
        name=values['Name']
        dob=values['DOB']
        age=values['Age']
        address=values['Address']
        udi=values['UDI']
        father_name=values['Father name']
        father_contact=values['Father contact']
        father_profession=values['Father profession']       
        father_aadhaar_no=values['Father aadhaar no']
        father_income=values['Father income']
        mother_name=values['Mother name']
        mother_contact=values['Mother contact']
        mother_profession=values['Mother profession']
        mother_aadhaar_no=values['Mother aadhaar no']
        mother_income=values['Mother income']    
        school=values['School']
        standard=values['Standard']
        academic_year=values['Academic Year']
        status=values['Status']
        medium=values['Medium']
        interests=values['Interests']
        disability=values['Disability']
        category=values['Category']
        try:
            summary_list="List added:"
            na="\nName:" +values['Name']
            summary_list+=na
            db="\nDOB:" +values['DOB']          
            summary_list+=db
            ag="\nAge:" +values['Age']
            summary_list+=ag
            add="\nAddress:" +values['Address']
            summary_list+=add
            ud="\nAge:" +values['UDI']
            summary_list+=ud
            fname="\nFather name:" +values['Father name']
            summary_list+=fname
            fcontact="\nFather contact:" +values['Father contact']
            summary_list+=fcontact
            fprof="\nFather profession:" +values['Father profession']   
            summary_list+=fprof
            faadhaar="\nFather aadhaar no:" +values['Father aadhaar no']
            summary_list+=faadhaar
            fincome="\nFather income:" +values['Father income']
            summary_list+=fincome
            mname="\nMother name:" +values['Mother name']
            summary_list+=mname
            mcontact="\nMother contact:" +values['Mother contact']
            summary_list+=mcontact
            mprof="\nMother profession:" +values['Mother profession']
            summary_list+=mprof
            maadhaar="\nMother aadhaar no:" +values['Mother aadhaar no']
            summary_list+=maadhaar
            mincome="\nMother income:" +values['Mather income']
            summary_list+=mincome
            sch="\nSchool:" +values['School']
            summary_list+=sch
            std="\nStandard:" +values['Standard']
            summary_list+=std
            ay="\nAcademic Year:" +values['Academic Year']
            summary_list+=ay
            stat="\nStatus:" +values['Status']
            summary_list+=stat
            med="\nMedium:" +values['Medium']
            summary_list+=med
            inte="\nInterests:" +values['Interests']
            summary_list+=inte
            disab="\nDisability:" +values['Disability']
            summary_list+=disab
            cat="\nCategory:" +values['Category']
            summary_list+=cat


             
            choice=sg.PopupOKCancel(summary_list,'Please confirm entry')


            if choice=='OK':
                print(event,values)
                df=pd.concat([df,pd.DataFrame([values])],ignore_index=True)
                df.to_excel(EXCEL_FILE,index=False)
                clear_inputs()
            else:
                sg.PopupOK('Edit entry')
        except:
            choice=sg.PopupOKCancel(summary_list,'Please confirm entry')
            sg.Popup('Data saved')
            print(event,values)
            df=pd.concat([df,pd.DataFrame([values])],ignore_index=True)
            df.to_excel(EXCEL_FILE,index=False)
            clear_inputs()


window.close()