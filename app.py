from flask import Flask,request,render_template,redirect,send_file
from flask_sqlalchemy import SQLAlchemy
from datetime import datetime
import os
from flask_mail import Mail,Message
import pandas as pd
import io

base_dir = os.path.abspath(os.path.dirname(__file__))
instance_dir = os.path.join(base_dir, 'instance')

app=Flask(__name__,instance_path=instance_dir)
app.config['SQLALCHEMY_DATABASE_URI']='sqlite:///Reg.db'
app.config.update(
    MAIL_SERVER='smtp.gmail.com',
    MAIL_PORT=587,
    MAIL_USERNAME=''

)
db=SQLAlchemy(app)

class Todo(db.Model):
    SNo=db.Column(db.Integer,primary_key=True)
    S1_name=db.Column(db.String(200),nullable=False)
    Clg_name=db.Column(db.String(200),nullable=False)
    Age=db.Column(db.String(200),nullable=False)
    Event_ID=db.Column(db.String(200),nullable=False)
    CDate=db.Column(db.DateTime,default=datetime.utcnow())


def OrmToList(reg_data):
    l=[]
    for i in reg_data:
        l.append([i.SNo,i.S1_name,i.Clg_name,i.Age,i.Event_ID,i.CDate])
    return l



@app.route('/',methods=['POST','GET'])
def index():
    if request.method=="POST":
        Sn=request.form['Name']
        College=request.form['Clg']
        SAge=request.form['Age']
        E_id=request.form['Eid']
        Reg_new=Todo(S1_name=Sn,Clg_name=College,Age=SAge,Event_ID=E_id)
        db.session.add(Reg_new)
        db.session.commit()
        return redirect('/registration')
    else:
        return render_template('index.html')
    
@app.route('/registration',methods=['POST','GET'])
def dlt():
    if request.method=="POST":
        return render_template("index.html")
    else:
        return render_template("RegSuccess.html")

@app.route('/admin/database')
def dba():
    query=db.session.query(Todo).all()
    return render_template('test.html', lists=query)

@app.route('/download7798')
def dwnld():
    query=db.session.query(Todo).all()

    data = [row.__dict__ for row in query]
    for row in data:
        row.pop('_sa_instance_state', None)  # Remove SQLAlchemy internal state

    # Create DataFrame
    df = pd.DataFrame(data)

    # Create Excel file in memory
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)

    output.seek(0)  # Move pointer to the start

    # Send file as download
    return send_file(
        output,
        as_attachment=True,
        download_name="Final_data.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    


if __name__=='__main__':
    app.run(debug=True)
