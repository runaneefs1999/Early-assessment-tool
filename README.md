# Early-assessment-tool
The early-assessment tool is a Python script that allows to take into account the environmental impact of a certain design based on a few basic design entries. Different structural materials are considered based on this input data.

The early-assessment tool tries to involve the environmental impact of structures as early as possible in a BIM workflow. This documentation provides a possible method to create a tool that based on simple inputs, generally available at the beginning of a design, gives an indication of the possible materials to be used for a bearing structure and its corresponding environmental impact.

In this particular tool a floor structure supported by underlying beams and columns is considered in different materials. The floor will be considered in cast-in-situ concrete, hollow-core slabs (HCS's) and cross-laminated timber (CLT). This floor in its turn will be supported by a frame structure consisting of two beams and four columns that will be considered in cast-in-situ concrete as well as steel and glued laminated timber. 
The floor and supporting frame structure under investigation in this tool are characterised by a certain span, height and specific load case. The general structure that can be studied with this tool is represented in the figure below.

![alt text](https://github.com/[runaneefs1999]/[Early-assessment tool]/blob/[branch]/image.jpg?raw=true)

To obtain the LCA scores of the possible bearing structures an LCA database was made in Excel, based on EcoInvent processes obtained through the SimaPro software. This database only includes LCA data for modules A1 to A3 as specified by EN 15978. This created LCA database is also included in the repository, as well as an additional Excel database used by the general python script to obtain correct material cross-sections.
