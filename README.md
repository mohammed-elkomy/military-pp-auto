# Egyption Army PowerPoint Automation

## Motivation
Due to excessive usage of MS-Office in my military service, I came up with a simple tool for automting about 90% of the tasks related to formating and printing. (Military-Style presentations, mostly in Arabic)

## Structure
Can be broken into 4 groups

1) ready add-in 
- This can be copied to MS-PowerPoint default add-in location then added to the toolbar

2. Actual Source presentation used for building the add-in in group [1]
- contains source.ppt the source presentation(the code is accompanied with the presentation) 
- config.komy.txt, heights.komy, toolbars which are files referenced by the code in source.ppt

3. macros folder contains the code exported from source.ppt (just to view the code on github)
4. gifs folder contains demo gifs shown in the readme

```
Numbers in brackets are related to the groups
.
├── [2] config.komy.txt     # config file for the tool
├── [4] gifs                # demo gifs
├── [2] heights.komy        # heights / number of lines data
├── [3] macros              # code exported as files
│   ├──  class modules      # class code files
│   ├──  forms              # form code files
│   └──  modules            # plain modules
├── [1] ready add-in        # the add-in to be added to MS-Powerpoint
├── [2] source.ppt          # source presentation with all exported code included
└── [2] toolbars            # toolbar images used by the tool (custom icons)
```
## Features
1. Slide Formatting
- Textbox colors (pre-defined set of colors)
- Text colors
- Basic Alignment and arrangement
- Line spacing
- fixing reversed numbers issues
- Detecting parenthesis and coloring text within 
- Optimal Alignment and line spacing (most advanced part and saving a lot of time)
<p align="center">
  <img src="https://raw.githubusercontent.com/mohammed-elkomy/military-pp-auto/master/gifs/1w.gif"  />
</p>
2. Presentation overall operations
- Black and White colorization for printers (since powerpoint black & white printing may miss some gradients and shadows causing high-ink consumption)
- Removing inner colors for color-printers (inner colors and shadows cause high-ink consumption)
<p align="center">
  <img src="https://raw.githubusercontent.com/mohammed-elkomy/military-pp-auto/master/gifs/2w.gif"  />
</p>
3. Other commonly used tools 

## Theory
What really makes me love this tool isn't only time-saving but the strong algorithmic and mathematical background needed to build those features
1. Optimal line spacing 
- Looks for the optimal line spacing for textboxes in a slide (they are all the same for visual symmetry) by defining a target padding percentage in a slide and try to minimize an euclidean distance function using ternary search algorithm (which makes it somehow real-time) 
- I also modelled the font sizes and number of lines in simple 3D tensor represented in heights.komy which makes searching for fonts and number of lines so efficient and apply only one UI call (not trial and error in the presentation view) [this is because MS-office wraps text so we can only find number of lines in reverse fashion (pre-processing approach)]
2. Finding horizontal groups of textboxes in one line
- I used some form of Jaccard index on the y-axis to find intersections (projected onto the y-axis)
- To handle rotations some trigonometry must be involved
<p align="center">
  <img src="https://github.com/mohammed-elkomy/military-pp-auto/raw/master/gifs/horizontal.png"  />
</p>
3. Handling nested groups
- I used recursion in the form of flood-fill through sub-groups

## Dedication
This work is dedicated for the Egyptian Army and I believe this helped a lot. 
