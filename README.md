# BOM Generator

## Installation

Install and activate the add-in [manually](https://knowledge.autodesk.com/support/fusion-360/troubleshooting/caas/sfdcarticles/sfdcarticles/How-to-install-an-ADD-IN-and-Script-in-Fusion-360.html) or using Jerome Briot's [Install scripts or addins from GitHub](https://apps.autodesk.com/FUSION/en/Detail/Index?id=789800822168335025&appLang=en&os=Win64&autostart=true) app.

## Usage

Open a drawing with a component heirarchy, navigate to the `Tools` panel of the `Design` workspace, and click the BOM Generator icon in the `Michael Greene` panel group.

When prompted, select a destination CSV file to export the BOM to.

The exported BOM will contain three columns (Part Number, Part Description and Quantity). Components which are actually subassemblies (those containing other components) are excluded from the export.
