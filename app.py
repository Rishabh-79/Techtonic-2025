from flask import Flask,request,render_template,redirect,send_file,flash
from flask_sqlalchemy import SQLAlchemy
from datetime import datetime,timezone,timedelta
import os
from flask_mail import Mail,Message
import pandas as pd
import io

base_dir = os.path.abspath(os.path.dirname(__file__))
instance_dir = os.path.join(base_dir, 'instance')
IST=timezone(timedelta(hours=5,minutes=30))

app=Flask(__name__,instance_path=instance_dir)
app.config['SQLALCHEMY_DATABASE_URI']='postgresql://registrationdb_1xph_user:4VVgNFIQc0TUQxlxJk0GM97TCkayQ7J3@dpg-d2fq81a4d50c73b53lf0-a/registrationdb_1xph'
app.config.update(
    MAIL_SERVER='smtp.gmail.com',
    MAIL_PORT=587,
    MAIL_USERNAME=''
)

db=SQLAlchemy(app)

class Todo(db.Model):
    SNo=db.Column(db.Integer,primary_key=True)
    S1_name=db.Column(db.String(200),nullable=False)
    Age1=db.Column(db.String(200),nullable=False)
    S2_name=db.Column(db.String(200),nullable=True)
    Age2=db.Column(db.String(200),nullable=True)
    Clg_name=db.Column(db.String(200),nullable=False)
    Event=db.Column(db.String(200),nullable=False)
    CDate=db.Column(db.DateTime,default=datetime.now(IST))


with app.app_context():
    db.create_all()


@app.route('/',methods=['POST','GET'])
def index():
    if request.method=="POST":
        Sn1=request.form['Name1']
        SAge1=request.form['Age1']
        Sn2=request.form['Name2']
        SAge2=request.form['Age2']
        College=request.form['Clg']
        SEvent=request.form['Event']
        Reg_new=Todo(S1_name=Sn1,Age1=SAge1,S2_name=Sn2,Age2=SAge2,Clg_name=College,Event=SEvent)
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
