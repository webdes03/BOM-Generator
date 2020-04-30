#Author-Michael Greene (mike-greene.com)
#Description-Export bill of materials as CSV

import adsk.core, adsk.fusion, adsk.cam, traceback

handlers = []

def run(context):
    ui = None
    try:
        app = adsk.core.Application.get()
        ui = app.userInterface
        workspace = ui.workspaces.itemById('FusionSolidEnvironment')
        tbPanels = workspace.toolbarPanels

        global tbPanel
        tbPanel = tbPanels.itemById('MichaelGreene')
        if tbPanel:
            tbPanel.deleteMe()
        tbPanel = tbPanels.add('MichaelGreene', 'Michael Greene', 'Michael Greene', False)

        cmdDefs = ui.commandDefinitions

        cmdDef = cmdDefs.itemById('N131MGBOM')
        if cmdDef:
            cmdDef.deleteMe()
        cmdDef = cmdDefs.addButtonDefinition('N131MGBOM', 'Generate BOM', 'Demo for new command', './/Resources//N131MGBOM')

        generateBomCmdCreated = GenerateBomCreatedEventHandler()
        cmdDef.commandCreated.add(generateBomCmdCreated)
        handlers.append(generateBomCmdCreated)


        button = tbPanel.controls.addCommand(cmdDef)
        button.isPromoted = True
        button.isPromotedByDefault = True

        #ui.messageBox('Hello addin')

    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))

class GenerateBomCreatedEventHandler(adsk.core.CommandCreatedEventHandler):
    def __init__(self):
        super().__init__()

    def notify(self, args):
        eventArgs = adsk.core.CommandCreatedEventArgs.cast(args)
        cmd = eventArgs.command

        # Connect to the execute event.
        onExecute = GenerateBomExecuteHandler()
        cmd.execute.add(onExecute)
        handlers.append(onExecute)

# Event handler for the execute event.
class GenerateBomExecuteHandler(adsk.core.CommandEventHandler):
    def __init__(self):
        super().__init__()        
    
    def notify(self, args):
        eventArgs = adsk.core.CommandEventArgs.cast(args)

        # Code to react to the event.
        app = adsk.core.Application.get()
        ui  = app.userInterface
        
        product = app.activeProduct
        design = adsk.fusion.Design.cast(product)
        if not design:
            ui.messageBox('No active design', 'Generate BOM')
            return
            
        ui.messageBox('Select a destination to output the Bill of Materials to...')
        
        # Get all occurrences in the root component of the active design
        root = design.rootComponent
        occs = root.allOccurrences
        
        # Gather information about each unique component
        bom = []
        for occ in occs:
            comp = occ.component
            jj = 0
            for bomI in bom:
                if bomI['component'] == comp:
                    # Increment the instance count of the existing row.
                    bomI['instances'] += 1
                    break
                jj += 1

            if jj == len(bom):
                # Gather any BOM worthy values from the component
                volume = 0
                bodies = comp.bRepBodies
                for bodyK in bodies:
                    if bodyK.isSolid:
                        volume += bodyK.volume
                
                # Add this component to the BOM
                if comp.occurrences.count <= 1:
                    bom.append({
                        'component': comp,
                        'partNumber': comp.partNumber,
                        'description': comp.description,
                        'instances': 1,
                        'volume': volume,
                        'childOccurrences': occ.childOccurrences.count,
                        'occurrences': comp.occurrences.count
                    })

        fileDialog = ui.createFileDialog()
        fileDialog.isMultiSelectEnabled = False
        fileDialog.title = "Set the file to save the BOM to..."
        fileDialog.filter = 'Text files (*.csv)'
        fileDialog.filterIndex = 0
        dialogResult = fileDialog.showSave()
            
        if dialogResult == adsk.core.DialogResults.DialogOK:
            filename = fileDialog.filename
        else:
            return
        
        result = "Part Number,Description,Quantity\n"
        for item in bom:
            result = result + "\"" + item['partNumber'] +  "\",\"" + item['description'] + "\",\"" + str(item['instances']) + "\"\n"
        
        with open(filename, "w") as outputFile:
            outputFile.writelines(result)

def stop(context):
    ui = None
    try:
        app = adsk.core.Application.get()
        ui  = app.userInterface
        #ui.messageBox('Stop addin')

        if tbPanel:
            tbPanel.deleteMe()

    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))
