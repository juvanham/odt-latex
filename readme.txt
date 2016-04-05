The aim of this small project is to help someone who has no idea about
LaTeX to create a text in openoffice and use that to generate LaTeX code.

The text that motivated this script was written in German and used sections, not chapters, however the program can be modified for this


The workflow I used is the following:

1) write a file odf.odt in the same directory and the convert script

2) run the convert script

3) now section_??.tex are generated in a tmp subdirectory.

4) I use a script to copy these to a directory with a main.tex file


5) this command is part of the main.tex file
\newcommand{\paragraphIndent}[1]{\vspace{4mm}\hspace{7mm}\begin{minipage}{149mm}#1\vspace{7mm}\end{minipage}}

6) I apply a set of patches to the resulting file: patch < section_01.patch

7) after fine tuning I regnerated these patches (diff -Naur ./tmp/section_01.tex section_01.tex >section_01.patch


Known limitations:
images are not supported, I used patches for this
citations with a page break are not converted


citations have two formats (author year)
 (Author Name year: remarks) this creates: ~\citep[remarks]{author-name:year}
 (Author Name o.J.: remarks) this creates: ~\citep[remakrs]{author-name:0000}

In the text can be (Abb. 4) this generates (Abb.~\ref{figure:4})

since colors were use as a help I did not implement different colors, but if a color is detects it can mark in orange {\color{orange} ... }

A way to improve this would be a table of hexrgb color and their name and look for the closed 3d distance to a specified color. At the end I needed to disable the color suport, which made this color approximation no longer needed.

The program is written in Ruby and it generates almost normal looking LaTeX code
by:
- breaking lines

- using references with lower case author names separated with dashes and after a colon the year.

- using LaTeX escapes







