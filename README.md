                                       
  Convert a PowerPoint (.pptx) file into a compilable LaTeX Beamer presentation.
                                                                                
  ## Requirements                                                               
                                                            
  - Python 3.10+
  - pip install python-pptx pillow
  - pdflatex (install via `brew install --cask mactex-no-gui`)                  
                                                                                
  ## Usage                                                                      
                                                                                
  **Step 1 — Convert your .pptx to .tex:**                                      
  cd scripts
  python3 pptx_to_beamer.py "../Your Slides.pptx"                               
                                                            
  **Step 2 — Verify the output:**                                               
  python3 verify_latex.py "Your Slides_beamer.tex"
                                                                                
  **Step 3 — Compile to PDF:**                              
  pdflatex "Your Slides_beamer.tex"                                             
                                                                                
  ## Options
                                                                                
  Choose a theme manually:                                  
  python3 pptx_to_beamer.py "../Your Slides.pptx" --theme Metropolis
                                                                                
  Available themes: Madrid, Warsaw, Berlin, Metropolis, AnnArbor, CambridgeUS,
  Boadilla, Hannover                                                            
  See `references/beamer-themes.md` for the full list.                          
                                                                                
  ## What it converts                                                           
                                                                                
  - Slide titles → `\begin{frame}{Title}`                                       
  - Bullet points → nested `\begin{itemize}` environments
  - Bold / italic / underline → `\textbf{}` / `\textit{}` / `\underline{}`      
  - Tables → `tabular` environments                                             
  - Images → extracted to `beamer_images/` and inserted with `\includegraphics` 
  - Free-floating text boxes → `\begin{block}{}` environments                   
                                                                                
  ## Files                                                                      
                                                                                
  pptx-to-beamer/                                           
  ├── SKILL.md                    # Agent skill instructions
  ├── scripts/                                                                  
  │   ├── pptx_to_beamer.py       # Main converter
  │   └── verify_latex.py         # Syntax checker                              
  ├── references/                                                               
  │   └── beamer-themes.md        # Theme reference guide
  └── assets/                                                                   
      └── beamer_template.tex     # Blank Beamer template   
                                                                                
