# wordle_vba
wordle remake with Excel's VBA.

This is my first coding project: a wordle game remake.
It's more akin to Termo at term.ooo, the portuguese version, since it is also in portuguese and I've used the same word list at https://github.com/fserb/pt-br repository.

As much as it would make a better, more readable code, I could not use modules. I've been having some unknown compatibility issues with my Excel or Windows version. For properties that can only be accessed with modules, I had to do some improvising.

Some of the .txt files are sketches and models I used. 

The main files are:

  "lista_palavras_semAcent_maisc.txt". The final word list after processing. All caps, no accents, no names.
  
  "nomes-registados-2017.csv". List used to remove names from the main list using power query.
  
  "ref". References.
  
  "OBJ". A crude roadmap and objectives list. The "ç" is completed and " * " is not implemented.
  
  "WORDLE". The sheet itself.
  
  "vba_raw". The main vba code inside "Planilha4(jogo)" sheet.
  
  "vba_core_raw". The main color checking logic.
  
  Before opening the workbook, I must remind you to enable macros and make the file trustworthy, otherwise Excel won't let the code run.
  
