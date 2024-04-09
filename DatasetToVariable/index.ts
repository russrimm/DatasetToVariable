import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { ComponentRenderer, IDatasetToExcelProps, IMakerStyleProps, IMakerButtonProps } from "./ComponentRenderer";
import * as React from "react";



export class DatasetToVariable implements ComponentFramework.ReactControl<IInputs, IOutputs> {
    private theComponent: ComponentFramework.ReactControl<IInputs, IOutputs>;
    private notifyOutputChanged: () => void;
    private _excelBase64: string;
    private readonly componentVersion: string = "1.0.0";

    /**
     * Empty constructor.
     */
    constructor() { }

    /**
     * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
     * Data-set values are not initialized here, use updateView.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
     * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
     * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
     */

    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary
    ): void {
        this.notifyOutputChanged = notifyOutputChanged;
        context.mode.trackContainerResize(true);
        // Assume a high pageSize to ensure all records are captured; adjust based on your needs
        context.parameters.DatasetToExport.paging.setPageSize(99999);
    }


    /**
     * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     * @returns ReactElement root react element for the control
     */
    public updateView(context: ComponentFramework.Context<IInputs>): React.ReactElement {

        const {
            TextColor, BGColor, IconColor, HoverTextColor, HoverBGColor, BorderColor,
            BorderHoverColor, BorderWidth, BorderRadius, Text, IconName, FileName, Loading
        } = context.parameters;

        // StylesProps object creation for the component
        const stylesProps: IMakerStyleProps = {
            textColor: TextColor.raw || "black",
            bgColor: BGColor.raw || "transparent", // Default to transparent if no color is provided
            iconColor: IconColor.raw || "inherit",
            hoverTextColor: HoverTextColor.raw || TextColor.raw || "black",
            hoverBgColor: HoverBGColor.raw || "transparent",
            borderColor: BorderColor.raw || "transparent",
            borderHoverColor: BorderHoverColor.raw || "transparent",
            borderWidth: BorderWidth.raw || 1,
            borderRadius: BorderRadius.raw || 0,
            buttonWidth: context.mode.allocatedWidth,
            buttonHeight: context.mode.allocatedHeight
        };

        // ButtonUIProps object creation for the component
        const buttonUiProps: IMakerButtonProps = {
            buttonText: Text.raw || "Export to Excel",
            iconName: IconName.raw || "Send" // Changed to 'Send' to better match the email sending context
        };

        // Props object creation for ComponentRenderer
        const props: IDatasetToExcelProps = {
            makerStyleProps: stylesProps,
            buttonProps: buttonUiProps,
            dataSet: context.parameters.DatasetToExport,
            selectedColumns: context.parameters.SelectedColumns,
            fileName: context.parameters.FileName.raw || `generated_file_${Date.now()}`,
            itemsLoading: context.parameters.DatasetToExport.loading,
            isLoading: Loading.raw || false, // Adjust based on your actual 'Loading' parameter
            // Add a placeholder for onButtonClick if there's no specific logic to execute
            onButtonClick: (event: React.MouseEvent<HTMLButtonElement>) => {
                // No operation (noop), or include any logic you want to execute on button click.
                console.log("Button clicked"); // Placeholder action
            },
            // Make sure to pass onBase64Ready if your ComponentRenderer expects it
            onBase64Ready: (base64String: string) => {
                this._excelBase64 = base64String; // Set the base64 string to your control's private variable
                console.log(base64String); // Keep the log if needed for debugging
                this.notifyOutputChanged(); // Notify the framework that the output has changed
            },
        };
    

        
       return React.createElement(ComponentRenderer, props);

    }

    /**
     * It is called by the framework prior to a control receiving new data.
     * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as "bound" or "output"
     */
    public getOutputs(): IOutputs {
        return { ExcelBase64: this._excelBase64 };
    }

    /**
     * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
     * i.e. cancelling any pending remote calls, removing listeners, etc.
     */
    public destroy(): void {
        // Add code to cleanup control if necessary
    }
}
