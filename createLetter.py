import pdfkit

def getLetter(name, teamName, position):
    imagePath= "letterDesign.png"

    htmlcontent = f'''
    <!DOCTYPE html>
    <html lang="en">
    <head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta http-equiv="X-UA-Compatible" content="ie=edge">
    <title>Letter Design</title>
    <style>
        body{{
            padding: 0;
            margin: 0;
        }}
        .container {{
            position: relative;
            text-align: center;
        }}

        .text {{
            color: rgba(1,41,64,255);
            position: absolute;
            top: 27%;
            height: 35%;
            padding: 0 14.2%;
            text-align: left;
            font-size: 1.6vw;
            font-family: Calibri, sans-serif;
            line-height: 1.4;
        }}
        img{{
            width: 100%;
            z-index: -1;
        }}
    </style>
    </head>
    <body>
    <div class="container">
        <img src="{imagePath}" alt="">
        <div class="text">
            <p>Dear {name}, <br><br></p>  
            <p>Welcome to the team of the ACM - Developers' Day 2024. We are excited to have you on board and look forward to the contributions you will make as <b>{position}</b> in Team <b>{teamName}</b>. <br><br></p>
            <p>It is worth acknowledging that the team of DevDay plays a crucial role in shaping the future of our event. Your involvement will be essential in crafting strategies and ensuring that our objectives align with the needs of our community. We believe in fostering a positive and collaborative work environment where everyone can grow and succeed together. <br><br> </p>
            <p>Our teams are committed to creating a supportive and inclusive environment where everyone's voices are heard and respected. We are confident that your addition to our team will further enhance our cohesion and effectiveness. <br><br></p>
            <p>Once again, welcome to the team! We are thrilled to have you with us and can't wait to see what we will achieve together.</p>     
        </div>
    </div>
    </body>
    </html>
    '''

    htmlFile = open("temp.html","w")     
    htmlFile.write(htmlcontent)  
    htmlFile.close()
    pdfkit.from_file("temp.html", "Appointment Letter.pdf")
    return "Appointment Letter.pdf"