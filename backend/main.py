from fastapi import FastAPI, UploadFile, File
from fastapi.middleware.cors import CORSMiddleware

app=FastAPI(title="Trimble Data Import AI")
app.add_middleware(CORSMiddleware,allow_origins=["https://jonasvanhecke-98.github.io","http://localhost:5500"],allow_credentials=True,allow_methods=["*"],allow_headers=["*"])

@app.get("/api/health")
def health():
    return {"ok":True}

@app.post("/api/analyze/local")
async def analyze_local(ifc:UploadFile=File(...),pdf:UploadFile=File(...)):
    return {"model_name":ifc.filename,"pdf_name":pdf.filename,"element_count":0,"assignment_count":0,"rows":[],"message":"Backend verbonden; volgende stap is IFC/OpenAI analyse."}

@app.post("/api/apply")
async def apply(payload:dict):
    rows=[r for r in payload.get("rows",[]) if r.get("status")=="confirmed"]
    return {"written":len(rows),"failed":0,"mode":"dry-run"}
