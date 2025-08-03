import adsk.core
import adsk.fusion
import os
import traceback
import subprocess
import platform
import re
from ...lib import fusionAddInUtils as futil
from ... import config

app = adsk.core.Application.get()
ui = app.userInterface


# TODO *** Specify the command identity information. ***
CMD_ID = f"{config.COMPANY_NAME}_{config.ADDIN_NAME}_exportAsSTL"
CMD_NAME = "Export bodies as STL"
CMD_Description = "A Fusion Add-in for exporting selected bodies as STL"

# Specify that the command will be promoted to the panel.
IS_PROMOTED = True

# TODO *** Define the location where the command button will be created. ***
# This is done by specifying the workspace, the tab, and the panel, and the
# command it will be inserted beside. Not providing the command to position it
# will insert it at the end.
WORKSPACE_ID = "FusionSolidEnvironment"
PANEL_ID = "SolidScriptsAddinsPanel"
COMMAND_BESIDE_ID = "ScriptsManagerCommand"

# Resource location for command icons, here we assume a sub folder in this directory named "resources".
ICON_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), "resources", "")

# Local list of event handlers used to maintain a reference so
# they are not released and garbage collected.
local_handlers = []


# Executed when add-in is run.
def start():
    # Create a command Definition.
    cmd_def = ui.commandDefinitions.addButtonDefinition(
        CMD_ID, CMD_NAME, CMD_Description, ICON_FOLDER
    )

    # Define an event handler for the command created event. It will be called when the button is clicked.
    futil.add_handler(cmd_def.commandCreated, command_created)

    # ******** Add a button into the UI so the user can run the command. ********
    # Get the target workspace the button will be created in.
    workspace = ui.workspaces.itemById(WORKSPACE_ID)

    # Get the panel the button will be created in.
    panel = workspace.toolbarPanels.itemById(PANEL_ID)

    # Create the button command control in the UI after the specified existing command.
    control = panel.controls.addCommand(cmd_def, COMMAND_BESIDE_ID, False)

    # Specify if the command is promoted to the main toolbar.
    control.isPromoted = IS_PROMOTED


# Executed when add-in is stopped.
def stop():
    # Get the various UI elements for this command
    workspace = ui.workspaces.itemById(WORKSPACE_ID)
    panel = workspace.toolbarPanels.itemById(PANEL_ID)
    command_control = panel.controls.itemById(CMD_ID)
    command_definition = ui.commandDefinitions.itemById(CMD_ID)

    # Delete the button command control
    if command_control:
        command_control.deleteMe()

    # Delete the command definition
    if command_definition:
        command_definition.deleteMe()


# Function that is called when a user clicks the corresponding button in the UI.
# This defines the contents of the command dialog and connects to the command related events.
def command_created(args: adsk.core.CommandCreatedEventArgs):
    # General logging for debug.
    futil.log(f"{CMD_NAME} Command Created Event")

    # https://help.autodesk.com/view/fusion360/ENU/?contextId=CommandInputs
    inputs = args.command.commandInputs

    # TODO Define the dialog for your command by adding different inputs to the command.

    lastUsedFolder = getLastUsedFolder()
    textBox = inputs.addTextBoxCommandInput(
        "folderPathInput", "File Path", lastUsedFolder, 1, True
    )

    browseButton = inputs.addBoolValueInput(
        "browseButton", "Browse...", False, "", True
    )

    selectionInput = inputs.addSelectionInput(
        "selectedBodies", "Select Bodies", "Select Bodies"
    )
    selectionInput.setSelectionLimits(0)
    selectionInput.addSelectionFilter("Bodies")

    filenameTable = inputs.addTableCommandInput('filenameTable', "Filenames", 1, "1")
    filenameTable.maximumVisibleRows = 10

    groupNameInput = inputs.addGroupCommandInput('groupNameInput', 'Advanced Naming')
    groupNameChildren = groupNameInput.children

    versionInput = groupNameChildren.addStringValueInput('versionInput', 'Version')

    prefixInput = groupNameChildren.addStringValueInput('prefixInput', 'Prefix')

    suffixInput = groupNameChildren.addStringValueInput('suffixInput', 'Suffix')

    nameFormatInput = groupNameChildren.addDropDownCommandInput('nameFormatInput', 'Name Formatting', 0)
    dropdownItems = nameFormatInput.listItems
    dropdownItems.add('None', True)
    dropdownItems.add('PascalCase', False)
    dropdownItems.add('camelCase', False)
    dropdownItems.add('snake_case', False)
    dropdownItems.add('kebab-case', False)
    dropdownItems.add('space separated', False)

    replaceButton = inputs.addBoolValueInput(
        "replaceButton", "Replace existing", True, "", False
    )
    replaceButton.tooltip = "Replace existing files"

    errorTextInput = inputs.addTextBoxCommandInput('errorTextInput', 'Log', '', 2, True)
    errorTextInput.isFullWidth = True

    # TODO Connect to the events that are needed by this command.
    futil.add_handler(
        args.command.execute, command_execute, local_handlers=local_handlers
    )
    futil.add_handler(
        args.command.inputChanged, command_input_changed, local_handlers=local_handlers
    )
    futil.add_handler(
        args.command.executePreview, command_preview, local_handlers=local_handlers
    )
    futil.add_handler(
        args.command.validateInputs,
        command_validate_input,
        local_handlers=local_handlers,
    )
    futil.add_handler(
        args.command.destroy, command_destroy, local_handlers=local_handlers
    )


# This event handler is called when the user clicks the OK button in the command dialog or
# is immediately called after the created event not command inputs were created for the dialog.
def command_execute(args: adsk.core.CommandEventArgs):
    # General logging for debug.
    futil.log(f"{CMD_NAME} Command Execute Event")

    # TODO ******************************** Your code here ********************************

    inputs = args.command.commandInputs

    selectionInput = inputs.itemById("selectedBodies")
    selectedFolder = inputs.itemById("folderPathInput").text
    filenameTable = inputs.itemById('filenameTable')
    replace = inputs.itemById("replaceButton").value

    exportSelectedBodies(selectionInput, selectedFolder, replace, filenameTable)


# This event handler is called when the command needs to compute a new preview in the graphics window.
def command_preview(args: adsk.core.CommandEventArgs):
    # General logging for debug.
    futil.log(f"{CMD_NAME} Command Preview Event")
    inputs = args.command.commandInputs

previous_selected_bodies = []

# This event handler is called when the user changes anything in the command dialog
# allowing you to modify values of other inputs based on that change.
def command_input_changed(args: adsk.core.InputChangedEventArgs):
    changed_input = args.input
    partial_inputs = args.inputs
    
    # Get the parent command object from the event arguments.
    command = args.firingEvent.sender
    # Access the complete collection of inputs from the command object.
    inputs = command.commandInputs

    # General logging for debug.
    futil.log(
        f"{CMD_NAME} Input Changed Event fired from a change to {changed_input.id}"
    )

    try:
        selectionInput = inputs.itemById("selectedBodies")
        selectedFolderInput = inputs.itemById("folderPathInput")
        filenameTable = inputs.itemById('filenameTable')
        versionNumberInput = inputs.itemById('versionInput')
        prefixInput = inputs.itemById('prefixInput')
        suffixInput = inputs.itemById('suffixInput')
        nameFormatInput = inputs.itemById('nameFormatInput')
    except:
        pass

    global previous_selected_bodies

    selected_bodies = []
    if not selectionInput:
        selected_bodies = previous_selected_bodies
    else:
        count = selectionInput.selectionCount
        for i in range(count):
            name = selectionInput.selection(i).entity
            selected_bodies.append(name)
        
    versionNumber = versionNumberInput.value
    prefix = prefixInput.value
    suffix = suffixInput.value
    formattingStyleIndex = nameFormatInput.selectedItem.index

    if changed_input.id == "browseButton":
        try:
            currentSelectedFolder = selectedFolderInput.text

            # Create filder dialog
            folderDialog = ui.createFolderDialog()
            folderDialog.title = "Select Destination Folder"
            folderDialog.initialDirectory = os.path.expanduser(currentSelectedFolder)

            # Show folder dialog
            dialogResult = folderDialog.showDialog()

            if dialogResult == adsk.core.DialogResults.DialogOK:
                # Update the text input with selected folder path
                path = folderDialog.folder
                pathInput = inputs.itemById("folderPathInput")
                pathInput.text = path

            # Reset button state
            changed_input.value = False
        except:
            futil.log("Failed:\n{}".format(traceback.format_exc()))

    elif changed_input.id in ['versionInput', 'prefixInput', 'suffixInput', 'nameFormatInput']:
        # When versionInput changes, update all filename textboxes in the table
        for i, body in enumerate(selected_bodies):
            body_name = body.name
            textBoxId = body.revisionId
            filename = generateFilename(body_name, versionNumber, prefix, suffix, formattingStyleIndex)
            nameInput = inputs.itemById(textBoxId)
            if nameInput:
                nameInput.value = filename

    elif changed_input.id == 'selectedBodies':
        added_body = [name for name in selected_bodies if name not in previous_selected_bodies]
        removed_body = [name for name in previous_selected_bodies if name not in selected_bodies]

        if added_body:
            for i, body in enumerate(added_body):
                # Create unique IDs and display names for each input
                body_name = body.name
                textBoxId = body.revisionId
                filename = generateFilename(body_name, versionNumber, prefix, suffix, formattingStyleIndex)
                textBoxValue = filename

                # Create a new TextBoxCommandInput
                subTextInput = inputs.addStringValueInput(textBoxId, '', textBoxValue)
                subTextInput.isFullWidth = True

                # Add to the next available row in the table
                row = filenameTable.rowCount
                filenameTable.addCommandInput(subTextInput, row, 0)
        elif removed_body:
            for i, body in enumerate(removed_body):
                try:
                    nameInput = inputs.itemById(body.revisionId)
                    (returnValue, row, column, rowSpan, columnSpan) = filenameTable.getPosition(nameInput)
                    delete_row = filenameTable.deleteRow(row)
                except:
                    futil.log("Failed:\n{}".format(traceback.format_exc()))

        # Update previous_selected_bodies for next change
        previous_selected_bodies = selected_bodies.copy()

    


# This event handler is called when the user interacts with any of the inputs in the dialog
# which allows you to verify that all of the inputs are valid and enables the OK button.
def command_validate_input(args: adsk.core.ValidateInputsEventArgs):
    # General logging for debug.
    futil.log(f"{CMD_NAME} Validate Input Event")

    inputs = args.inputs

    # Verify the validity of the input values. This controls if the OK button is enabled or not.

    # Get the selected folder path
    folderPathInput = inputs.itemById("folderPathInput")
    selectedFolder = folderPathInput.text
    selectionInput = inputs.itemById("selectedBodies")
    filenameTable = inputs.itemById('filenameTable')

    errorTextInput = inputs.itemById('errorTextInput')

    # Validate selected folder
    if not selectedFolder or not os.path.exists(selectedFolder):
        args.areInputsValid = False
        errorTextInput.text = "Selected folder does not exist."
        return
    else:
        errorTextInput.text = ''

    # Validate selection
    if selectionInput.selectionCount == 0:
        args.areInputsValid = False
        errorTextInput.text = "No bodies selected."
        return
    else:
        errorTextInput.text = ''

    # Validate filenames
    if filenameTable.rowCount:
        for i in range(filenameTable.rowCount):
            nameInput = filenameTable.getInputAtPosition(i, 0)
            name = nameInput.value

            # Check if name contains any invalid characters
            invalid_pattern = r'[<>:"/\\|?*\x00-\x1F]'
            if re.search(invalid_pattern, name):
                args.areInputsValid = False
                errorTextInput.text = f"Filename '{name}' contains invalid characters."
                return
    
    # Check for repeated filenames
    filenames = []
    for i in range(filenameTable.rowCount):
        nameInput = filenameTable.getInputAtPosition(i, 0)
        name = nameInput.value
        filenames.append(name)
    duplicates = set([x for x in filenames if filenames.count(x) > 1])
    if duplicates:
        args.areInputsValid = False
        errorTextInput.text = f"Duplicate filenames found: {', '.join(duplicates)}"
        return
        



# This event handler is called when the command terminates.
def command_destroy(args: adsk.core.CommandEventArgs):
    # General logging for debug.
    futil.log(f"{CMD_NAME} Command Destroy Event")

    global local_handlers, previous_selected_bodies
    previous_selected_bodies = []
    local_handlers = []


def exportSelectedBodies(selectionInput, exportFolder, replace, filenameTable):
    try:
        design = app.activeProduct

        if not design:
            futil.log("No active design found.")
            return

        # Filter selected bodies
        selectedBodies = []
        count = selectionInput.selectionCount
        for i in range(count):
            entity = selectionInput.selection(i).entity
            if entity.objectType == adsk.fusion.BRepBody.classType():
                selectedBodies.append(entity)

        # Export each selected body
        exportMgr = design.exportManager
        successCount = 0
        exportedFiles = []

        for i, body in enumerate(selectedBodies):
            try:
                # Create STL export options
                stlOptions = exportMgr.createSTLExportOptions(body)
                
                filenameInput = filenameTable.getInputAtPosition(i, 0)
                fileName = filenameInput.value
                filePath = os.path.join(exportFolder, fileName)

                if not replace:
                    counter = 1
                    while os.path.exists(filePath):
                        name, ext = os.path.splitext(fileName)
                        name = re.sub(r'\(\d+\)', '', name)
                        fileName = f"{name}({counter}){ext}"
                        filePath = os.path.join(exportFolder, fileName)
                        counter += 1

                # Set export options
                stlOptions.filename = filePath
                stlOptions.meshRefinement = (
                    adsk.fusion.MeshRefinementSettings.MeshRefinementMedium
                )
                stlOptions.isBinaryFormat = True  # Binary STL is more compact

                # Export the body
                exportMgr.execute(stlOptions)
                successCount += 1
                exportedFiles.append(fileName)

            except Exception as e:
                ui.messageBox(f'Failed to export body "{body.name}": {str(e)}')
                futil.log("Failed to export:\n{}".format(traceback.format_exc()))

        # Show completion message
        if successCount > 0:
            fileList = "\n".join(f"• {file}" for file in exportedFiles)
            text_message = f"Successfully exported {successCount} of {len(selectedBodies)} bodies to:\n{exportFolder}\n\nFiles created:\n{fileList}"
            returnValue = ui.messageBox(text_message, 'Open location?', 3)

            if returnValue == 2:
                openFolderLocation(exportFolder)

            saveLastUsedFolder(exportFolder)

        else:
            futil.log("No bodies were exported successfully.")

    except:
        futil.log("Failed to export:\n{}".format(traceback.format_exc()))

def generateFilename(bodyName, versionNumber, prefix, suffix, formattingStyleIndex):
    name = sanitize_filename(bodyName)
    versionStr = str(versionNumber).strip()

    nameSplit, ext = os.path.splitext(name)
    
    # Replace hyphens, underscores, and spaces with a single space
    nameSplit = re.sub(r"[-_ ]+", " ", nameSplit)

    # Insert space before uppercase letters that follow lowercase (e.g., fileName → file Name)
    nameSplit = re.sub(r"([a-z])([A-Z])", r"\1 \2", nameSplit)

    # Split into words and lowercase them all
    words = nameSplit.lower().split()

    if formattingStyleIndex == 1: #PascalCase
        name = ''.join(w.capitalize() for w in words) + ext
        versionStr = f"V{versionStr}" if versionStr else ""
    elif formattingStyleIndex == 2: #camelCase
        name = words[0] + ''.join(w.capitalize() for w in words[1:]) + ext
        versionStr = f"V{versionStr}" if versionStr else ""
    elif formattingStyleIndex == 3: #snake_case
        name = '_'.join(words) + ext
        versionStr = f"_v{versionStr}" if versionStr else ""
    elif formattingStyleIndex == 4: #kebab-case
        name = '-'.join(words) + ext
        versionStr = f"-v{versionStr}" if versionStr else ""
    elif formattingStyleIndex == 5: #space separated
        name = ' '.join(words) + ext
        versionStr = f" v{versionStr}" if versionStr else ""
    else:
        pass

    return f'{prefix}{name}{versionStr}{suffix}.stl'

def getLastUsedFolder():
    # Get the last used folder from document attributes
    try:
        app = adsk.core.Application.get()
        doc = app.activeDocument

        if doc:
            # Look for our stored attribute
            attrib = doc.attributes.itemByName("ExportTools", "LastUsedFolder")
            if attrib:
                folderPath = attrib.value
                if os.path.exists(folderPath):
                    return folderPath

        # Default to user's Documents folder if no last used folder
        return os.path.expanduser("~")

    except:
        return os.path.expanduser("~")


def saveLastUsedFolder(folderPath):
    # Save the folder path to document attributes
    futil.log('Start saving folder path!')
    try:
        app = adsk.core.Application.get()
        doc = app.activeDocument

        if doc:
            # Remove existing attribute if it exists
            attrib = doc.attributes.itemByName("ExportTools", "LastUsedFolder")
            if attrib:
                attrib.deleteMe()

            # Add new attribute with folder path
            doc.attributes.add("ExportTools", "LastUsedFolder", folderPath)
        
        futil.log('Folder path save successfully!')
    except:
        futil.log("Failed:\n{}".format(traceback.format_exc()))


def openFolderLocation(folderPath):
    # Open the folder location in the system's file manager with better error handling
    try:
        # Ensure the folder exists
        if not os.path.exists(folderPath):
            app = adsk.core.Application.get()
            ui = app.userInterface
            futil.log(f"Folder does not exist:\n{folderPath}")
            return

        system = platform.system()

        if system == "Windows":
            # Windows - open in Explorer and select the folder
            subprocess.run(["explorer", "/select,", folderPath], check=True)
        elif system == "Darwin":  # macOS
            # macOS - open in Finder
            subprocess.run(["open", folderPath], check=True)
        elif system == "Linux":
            # Linux - try common file managers
            try:
                subprocess.run(["nautilus", folderPath], check=True)
            except subprocess.CalledProcessError:
                try:
                    subprocess.run(["dolphin", folderPath], check=True)
                except subprocess.CalledProcessError:
                    subprocess.run(["xdg-open", folderPath], check=True)
        else:
            # Fallback
            subprocess.run(["open", folderPath], check=False)

    except Exception as e:
        # If opening fails, show the path in a message
        app = adsk.core.Application.get()
        ui = app.userInterface
        ui.messageBox(
            f"Could not open folder automatically.\nFiles saved to:\n{folderPath}"
        )

def sanitize_filename(name):
    # Remove invalid characters
    name = re.sub(r'[<>:"/\\|?*\x00-\x1F]', '', name)
    # Remove trailing spaces or dots
    return name.rstrip(" .")