from flask import Flask,render_template,request,session,redirect, url_for
import os
import shutil
from ms_teams import add_team_func


app=Flask(__name__)

@app.route('/',methods=['GET','POST'])
def home():

    if request.method=='POST':
        email=request.form.get('email')
        pas=request.form.get('pas')
        team_name=request.form.get('team name')
        col_name=request.form.get('col name')
        file=request.files['file']

        if os.path.isdir('file'):
            shutil.rmtree('file')

        os.mkdir('file')
        file_name=file.filename
        file_name='file/'+file_name
        file.save(file_name)


        import pandas as pd
        paths='file/'+file.filename
        try:
            if file.filename.split('.')[1]=='xlsx':
                df=pd.read_excel(paths,usecols=[col_name])
            elif file.filename.split('.')[1]=='csv':
                df=pd.read_csv(paths,usecols=[col_name])
        except:
            return render_template('error.html',msg='Invalid Column Name')

        df=df.loc[df[col_name].str.contains('@')]
        if len(df)==0:
            return render_template('error.html',msg='Invalid Excel File or Column Name')
        df.loc[len(df)]=''


        msg=add_team_func(email,pas,team_name,file.filename,col_name)
        print(msg)
        if msg!='':
            return render_template('error.html',msg=msg)
        
        return render_template('success.html')
    
    return render_template('home.html')

# @app.route('/success')
# def success():
#     return render_template()



app.run(debug=True)