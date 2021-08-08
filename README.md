# Python Email sender with Excel Data Manipulation

###### Version 1.0

Python script to automate the sending of emails with data that is read from excel files. Useful for teachers and professor who work with student notes, and other professionals with similar requirements.

### Features

- Works with gmail accounts
- Send email messages in HTML markup
- Gmail account credentials are requested in the execution of the script
- The data is read from excel files (only works with xlsx format)

### Installation & run

Create a Python3 virtual environment and activate it:

```python
virtualenv venv
source venv/bin/activate
```

Clone the GitHub repo:

```git
git@github.com:Joh4nnSmith/email-sender.git
```

Install Python dependencies:

```python
cd email-sender
pip3 install -r requirements.txt
```

Open the `email_sender.py` file and edit the html section of the email body:

> **self.receiver_name -** Name read from excel file, do not modify
>
> **data[0], data[1], data[N]** - Organize according to email structure

```python
# Create_message function of the Email class
# Example with sending grades to students

html_body = f'''
     	<html>
         	<head> </head>
         	<body>
                Good morning, {self.receiver_name}

                <p> Below are the programming course notes. The global grade is divided 				as follows: Examen (50%) y Workshop (50%). <br> <br>

                In general, your notes are: <br>
                <ul>
                    <li type="disc"><b>Examen</b>: {data[1]}</li>
                    <li type="disc"><b>Workshop</b>: {data[2]}</li>
                </ul>

                The total grade for the course is: {data[0]} <br> <br>
                <br> <br>
                </p
      		</body>
   		</html>
        '''
```

Save changes and run the script:

```python
`python3 email-sender.py`
```













