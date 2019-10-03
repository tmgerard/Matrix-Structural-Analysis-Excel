# Excel Matrix Structural Analysis
The purpose of this project is to learn techniques of solving linear systems and create an object-oriented matrix structural analysis program within Microsoft Excel using VBA and to gain a better understanding of how commercial analysis software works. Excel was chosen as the platform for this project because many structural engineers use it in practice.
## Goal
The ultimate goal is to create a simple structural analysis engine that uses the direct stiffness method to solve plane and space trusses and frames, that can be extended through proper use of object-oriented concepts.
## Desired Features
* Matrix and vector classes (Dense and Sparse)
* Simultaneous equation algorithms (Gaussian Elimination and Cholesky Decomposition)
* Element classes for defining member stiffness matrices
* Structural material library (AASHTO and ASTM)
* Rolled structural shape library (AISC shapes)
* Global stiffness matrix assembly algorithm
* Structural system input
* Analysis output
## VBA Editor Add-In
The VBA editor is notorious for being out of date with its features and not user friendly for larger scale projects with many modules and class modules. The [Rubberduck VBA Add-In](https://github.com/rubberduck-vba/Rubberduck) is recommended to aid in module organization and is required to run the units tests in this project.
## References
* [Newton Excel Bach](https://newtonexcelbach.com/) - Great reference for structural engineering spreadsheets.
* [Rubberduck VBA Blog](https://rubberduckvba.wordpress.com/) - The Rubberduck VBA blog has some great entries regarding object-oriented design using VBA.
* [Matrix Structural Analysis, 2nd Edition \[McGuire, Gallagher, and Ziemian\]](http://www.mastan2.com/textbook.html) - Textbook available through the MASTAN2 website as a free PDF download.
* [MASTAN2](http://www.mastan2.com/) - Free structural analysis program written in MATLAB that is based on the analysis as presented in the Matrix Structural Analysis, 2nd Edition text.
* [Finite Element Procedures, 2nd Edition \[Bathe\]](http://www.adina.com/pubs/publications50.shtml) - Textbook on the finite element method that has details for the implementation of global stiffness matrix assembly. Available as a free PDF download.