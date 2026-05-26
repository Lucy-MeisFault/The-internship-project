from fastapi import FastAPI, UploadFile, File
import pandas as pd
import io
from pptx import Presentation
app = FastAPI()


from fastapi.middleware.cors import CORSMiddleware

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


    



@app.post("/upload")
async def upload(file: UploadFile = File(...)):
    contents = await file.read()


    if file.filename.endswith('.csv'):
        df = pd.read_csv(io.BytesIO(contents))
    else:
        df = pd.read_excel(io.BytesIO(contents)) 
    lines = {}
    for line in df['Jakou linku provozované společností ARRIVA vlaky s.r.o., využívám nejčastěji?'].unique():
        lines[line] = df[df['Jakou linku provozované společností ARRIVA vlaky s.r.o., využívám nejčastěji?'] == line]

    presentation = Presentation();



        

    # aight here i'd prooobably make methods for putting the data into whatever chart type they had in the presentation
    # basically a thingy that takes data on sometbuhing makes graph and creates a slide with the title of idk whatever the category is 
    # and that for like each type of those? in case he'd want me to make it actually modular and usable for other things than the exact thing here also i could maybe put it on me github
    # and yeah then i'd probably write a niceee for loop where i'd go over the aforementioned list and basically make said slide for line 0 1 2 3 etc
    # and then we'd go onto a different category? that would be how i'd make it if he doesnt want it fancy 
    # if he does i'll probably somehow make the jscript hand over like a dictionary or something with the exact instructions like first we'll have a pie chart of data x for all the lines
    # which should be easier to do when i have the methods i mentioned before
    # and then we ship it? i think? 
    return {"filename": file.filename}


