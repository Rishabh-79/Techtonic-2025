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

app.config['MAIL_SERVER'] = 'smtp.gmail.com'
app.config['MAIL_PORT'] = 587
app.config['MAIL_USE_TLS'] = True
app.config['MAIL_USE_SSL'] = False
app.config['MAIL_USERNAME'] = '59rishabh@gmail.com'       
app.config['MAIL_PASSWORD'] = 'wrsd jnsk dfgl qmpd '         
app.config['MAIL_DEFAULT_SENDER'] = '59rishabh@gmail.com'

mail=Mail(app)

db=SQLAlchemy(app)

class Todo(db.Model):
    SNo=db.Column(db.Integer,primary_key=True)
    S1_name=db.Column(db.String(200),nullable=False)
    S2_name=db.Column(db.String(200),nullable=True)
    S3_name=db.Column(db.String(200),nullable=True)
    S4_name=db.Column(db.String(200),nullable=True)
    S5_name=db.Column(db.String(200),nullable=True)
    Mailid=db.Column(db.String(200),nullable=False)
    Clg_name=db.Column(db.String(200),nullable=False)
    Department=db.Column(db.String(200),nullable=False)
    Event=db.Column(db.String(200),nullable=False)
    CDate=db.Column(db.DateTime,default=datetime.now(IST))


with app.app_context():
    db.create_all()


@app.route('/',methods=['POST','GET'])
def index():
    return render_template('index.html')
    

@app.route('/registration',methods=['POST','GET'])
def dlt():
    if request.method=="POST":
        return render_template("index.html")
    else:
        return render_template("RegSuccess.html")

# Route to view data
@app.route('/admin/database')
def dba():
    query=db.session.query(Todo).all()
    return render_template('test.html', lists=query)

# Route to download data as an excel file
@app.route('/download7798')
def dwnld():
    query = db.session.query(Todo).all()

    # Extract data
    data = [row.__dict__ for row in query]
    for row in data:
        row.pop('_sa_instance_state', None)  # Remove SQLAlchemy internal state

    # Define the column order as per model definition
    columns_order = ['SNo', 'S1_name', 'S2_name', 'S3_name', 'S4_name', 'S5_name',
                     'Mailid', 'Clg_name', 'Department', 'Event', 'CDate']

    # Create DataFrame with specified column order
    df = pd.DataFrame(data, columns=columns_order)

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

    
@app.route("/registerall",methods=['POST','GET'])
def registerall():
    try:
        if request.method=="POST":
            Sn1=request.form['Name1']
            Sn2=request.form['Name2']
            Sn3=''
            Sn4=''
            Sn5=''
            mailval=request.form['mailid']
            College=request.form['Clg']
            Sdept=request.form['Dept']
            SEvent=request.form['Event']
            tor=datetime.now(IST)
            Reg_new=Todo(S1_name=Sn1,S2_name=Sn2,S3_name=Sn3,S4_name=Sn4,S5_name=Sn5,Mailid=mailval,Department=Sdept,Clg_name=College,Event=SEvent,CDate=tor)
            db.session.add(Reg_new)
            db.session.commit()
            msg= Message(subject="Registration for Techtonic 2025",
            recipients=[mailval],  # list of recipient emails
            body='''Dear Team,

            Thank you for registering for Techtonic 2025, the Inter-Collegiate Technical Fest of the Department of Computer Science, Madras Christian College.

            We are pleased to confirm your registration for the event. Kindly note the following:

            Date: Friday, 12th September 2025

            Offline Registration: 9:00 AM at the venue

            Inauguration Ceremony: 9:30 AM at Anderson Hall

            Please ensure that all team members carry their College ID card and a Bonafide Certificate signed by the Head of the Department.

            We look forward to your participation and wish you the very best.

            Warm regards,
            Organizing Team – Techtonic 2025
            Department of Computer Science
            Madras Christian College''')
            mail.send(msg)
            return redirect('/registration')
        else:
            return render_template("RegformOthers.html")
    except:
        return render_template("Failure.html")

@app.route("/registerazp",methods=['POST','GET'])
def registerazp():
    try:
        if request.method=="POST":
            Sn1=request.form['Name1']
            Sn2=request.form['Name2']
            Sn3=request.form['Name3']
            Sn4=request.form['Name4']
            Sn5=request.form['Name5']
            mailval=request.form['mailid']
            College=request.form['Clg']
            Sdept=request.form['Dept']
            SEvent=request.form['Event']
            Reg_new=Todo(S1_name=Sn1,S2_name=Sn2,S3_name=Sn3,S4_name=Sn4,S5_name=Sn5,Mailid=mailval,Department=Sdept,Clg_name=College,Event=SEvent)
            db.session.add(Reg_new)
            db.session.commit()
            msg= Message(subject="Registration for Techtonic 2025",
            recipients=[mailval],  # list of recipient emails
            body='''Dear Team,

            Thank you for registering for Techtonic 2025, the Inter-Collegiate Technical Fest of the Department of Computer Science, Madras Christian College.

            We are pleased to confirm your registration for the event. Kindly note the following:

            Date: Friday, 12th September 2025

            Offline Registration: 9:00 AM at the venue

            Inauguration Ceremony: 9:30 AM at Anderson Hall

            Please ensure that all team members carry their College ID card and a Bonafide Certificate signed by the Head of the Department.

            We look forward to your participation and wish you the very best.

            Warm regards,
            Organizing Team – Techtonic 2025
            Department of Computer Science
            Madras Christian College''')
            mail.send(msg)
            return redirect('/registration')
        else:
            return render_template("RegformAzp.html")
    except:
        return render_template("Failure.html")

if __name__=='__main__':
    app.run(debug=False)
