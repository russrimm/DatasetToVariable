import * as React from "react";
import { IIconProps, IButtonStyles, DefaultButton, Spinner } from "@fluentui/react";
import * as XLSX from 'xlsx';

export interface IMakerStyleProps {
    textColor: string;
    bgColor: string;
    iconColor: string;
    hoverTextColor: string;
    hoverBgColor: string;
    borderColor: string;
    borderHoverColor: string;
    borderWidth: number;
    borderRadius: number;
    buttonWidth: number;
    buttonHeight: number;
}

export interface IMakerButtonProps {
    buttonText: string;
    iconName: string;
}

export interface IDatasetToExcelProps {
    makerStyleProps: IMakerStyleProps;
    buttonProps: IMakerButtonProps;
    dataSet: ComponentFramework.PropertyTypes.DataSet;
    selectedColumns: ComponentFramework.PropertyTypes.DataSet;
    fileName: string;
    itemsLoading: boolean;
    isLoading: boolean;
    onButtonClick: (event: React.MouseEvent<HTMLButtonElement>) => void;
    onBase64Ready?: (base64: string) => void; // Callback for base64 string
}

export const ComponentRenderer = (props: IDatasetToExcelProps) => {
    const { makerStyleProps, buttonProps, dataSet, selectedColumns, fileName, itemsLoading, isLoading, onBase64Ready } = props;
    const buttonIcon: IIconProps = { iconName: buttonProps.iconName };

    const handleClick = (event: React.MouseEvent<HTMLButtonElement>) => {
      props.onButtonClick(event);
      console.log("Total Records: ", dataSet.paging.totalResultCount);
      // Temporarily commented out const dataToExport = prepareData(dataSet, selectedColumns);
      const dataToExport = [{ Name: "John Doe", Email: "john.doe@example.com" }, { Name: "Jane Doe", Email: "jane.doe@example.com" }];
      console.log("Data to export:", dataToExport);
      // If dataToExport is empty, then there is no data to export
      if (dataToExport.length === 0) {
          console.log("No data to export");
          return;
        }
        const worksheet: XLSX.WorkSheet = XLSX.utils.json_to_sheet(dataToExport);
        const workbook: XLSX.WorkBook = { Sheets: { 'data': worksheet }, SheetNames: ['data'] };

        //Use XLSX.write() to get the workbook as a base64 string
        const base64string = XLSX.write(workbook, { bookType: 'xlsx', type: 'base64' });
        console.log("Base64 String:", base64string);
        if (props.onBase64Ready) {
            props.onBase64Ready(base64string);
        }
    };

    const isLoadingState = itemsLoading || isLoading;

    return (
        <DefaultButton
            styles={getStyle(makerStyleProps)}
            title="Send Attachment"
            ariaLabel="Send Attachment"
            disabled={false}
            onClick={handleClick}
            iconProps={{ iconName: "Mail" }}
        >
           
        </DefaultButton>
    );
};

interface TypedEntityRecord {
  getValue(columnName: string): any;
}

// Define interfaces to describe the structure more clearly to TypeScript.

interface IColumnValue {
  displayName?: string;
}

interface IRecord {
  [key: string]: any; // This allows indexing with a string to return any type.
}

interface IColumns {
  [columnName: string]: IColumnValue;
}

// Update the function to use these interfaces for better type checking.

const getColumnNames = (dataSet: ComponentFramework.PropertyTypes.DataSet): Array<{ key: string, value: string }> => {
    const columns: IColumns = dataSet.columns as unknown as IColumns; // Use 'as unknown as' for a two-step assertion if direct casting is problematic.
    return Object.keys(columns).map(key => {
        const column: IColumnValue = columns[key]; // Now TypeScript knows what to expect when accessing columns[key].
        return { key, value: column.displayName || key };
    });
};

const prepareData = (
  dataSet: ComponentFramework.PropertyTypes.DataSet, 
  selectedColumns: ComponentFramework.PropertyTypes.DataSet | null = null
): any[] => {
    const columnList = getColumnNames(selectedColumns || dataSet);
    return dataSet.sortedRecordIds.map(recordId => {
        const record: IRecord = dataSet.records[recordId] as unknown as IRecord; // Assuming records can be indexed with a string key.
        let rowData: { [key: string]: any } = {};
        columnList.forEach(({ key }) => {
            rowData[key] = record[key]; // This access pattern should be clear to TypeScript now.
        });
        return rowData;
    });
};



const getStyle = (styleProps: IMakerStyleProps): IButtonStyles => {
    const borderStyle = styleProps.borderWidth && styleProps.borderWidth > 0 ? `solid ${styleProps.borderWidth}px ${styleProps.borderColor}` : "none";
    return {
        root: {
            color: styleProps.textColor,
            backgroundColor: styleProps.bgColor,
            border: borderStyle,
            borderRadius: `${styleProps.borderRadius}px`,
            width: `${styleProps.buttonWidth}px`,
            height: `${styleProps.buttonHeight}px`,
        },
        icon: {
            color: styleProps.iconColor,
        },
        rootHovered: {
            color: styleProps.hoverTextColor,
            backgroundColor: styleProps.hoverBgColor,
            border: `solid ${styleProps.borderWidth}px ${styleProps.borderHoverColor}`,
        },
    };
};
