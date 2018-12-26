# Convert Google Sheets table to LaTeX

Small script to convert the selected area from a Google Sheets document to a LaTeX table.

What's supported:
- Merged columns
- Merged rows
- Multi-row cells (newline within a cell)
- Frequently used special characters

# Installation

You can install it on the Sheets you are using on, it's not a global add-on, like the ones in the store.

- Go to tools, Script Editor.
- Paste the content of the .js file into the editor.
- Save, close the editor
- Select the range you want to convert
- Click Add-ons / LaTeX Table / Selection to LaTeX table

# Sample

## In Google Sheets

![gsheets_input](https://user-images.githubusercontent.com/3745745/50458420-0a6b9a80-09b7-11e9-92f4-61dfdeaf212b.png)

## Generated LaTeX code

```
\begin{table}
\centering
\begin{tabular}{|l|l|l|l|l|}
\hline
& & \multicolumn{2}{c|}{multi column} & \\
\hline
& & col1 & col2 & col3 \\
\hline
& row1 & data1 & data4 & data7 \\
\hline
\multirow{2}{*}{multi row} & row2 & data2 & \begin{tabular}[c]{@{}l@{}}data5\\multiline cell\end{tabular} & data8 \\
\cline{2-5}
& row3 & data3 & symbols work too \$ \textasciitilde \textbackslash \# & data9 \\
\hline
\end{tabular}
\label{table:table}
\caption{\small{}} 
\end{table}
```

## LaTeX visual

![latex_visual](https://user-images.githubusercontent.com/3745745/50458421-0a6b9a80-09b7-11e9-8557-e6eb334c2057.png)
