# DIQB-UC-HPLC
Read PDF files generated by HPLC (High Performance Liquid Chromatography) instruments and reinterpret them as Excel spreadsheets _.xlsx_ files.

Scripts written for an iPRE (Investigación en Pregrado UC) program I took part in second university year (2018).
The research program was guided by Eduardo Agosin, Professor at DIQB (Departamento de Ingeniería Química y Bioprocesos), Escuela de Ingeniería, Pontificia Universidad Católica de Chile.

### Current features
- Reads standard curves and fits sample areas to estimate their concentrations
- Generates plots for standard curves, as well as linear fit parameters
- Supports multiple molecule measurements and concentration reporting, as long as they are labeled accordingly
- Supports internal standards
- Supports glucose and lipid-related measurements and spreadsheet generation

The script is written entirely in `Python 3.6` and its requirements are included in `requirements.txt` file

### Upcoming features
- Compatibility with raw signal data obtained from HPLC for plotting and further analysis

