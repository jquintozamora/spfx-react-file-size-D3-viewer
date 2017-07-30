import * as React from "react";

import styles from "./FileSizeD3Viewer.module.scss";

// import PnP JS Core
import pnp from "sp-pnp-js";

// import models
import { MyDocument, MyFolder, MyFolderCollection, ITreeMapNode } from "../model";

// import custom parsers
import { SelectDecoratorsArrayParser } from "../parser/SelectDecoratorsParsers";


import TreeMap from "react-d3-treemap";
// include its styles in you build process as well
import "react-d3-treemap/dist/react.d3.treemap.css";
import ContainerDimensions from "react-container-dimensions";

// import React props and state
import { IFileSizeD3ViewerProps } from "./IFileSizeD3ViewerProps";
import { IFileSizeD3ViewerState } from "./IFileSizeD3ViewerState";

export default class FileSizeD3Viewer extends React.Component<IFileSizeD3ViewerProps, IFileSizeD3ViewerState> {
  constructor(props: IFileSizeD3ViewerProps) {
    super(props);
    // set initial state
    this.state = {
      data: null,
      errors: []
    };
  }

  public render(): React.ReactElement<IFileSizeD3ViewerProps> {
    return (
      <div>
        {this._getUIElement()}
        <div>
          {
            this.state.errors.length > 0
              ? this.state.errors.map(item => <div>{item.toString()}</div>)
              : null
          }
        </div>
      </div>
    );
  }

  public componentDidMount(): void {
    const libraryName: string = "Documents";
    console.log("libraryName: " + libraryName);
    this._readAllFilesSize(libraryName);
  }

  private _getUIElement() {
    return this.state.data !== null ?
      <ContainerDimensions>
        {({ width, height }) =>
          <TreeMap
            width={width - 20}
            height={350}
            data={this.state.data}
            valueUnit={"KB"}
          />
        }
      </ContainerDimensions>
      :
      <div>Loading...</div>;
  }

  private async _readAllFilesSize(libraryName: string): Promise<void> {
    try {

      let docsInTreeMap: ITreeMapNode[] = [];

      // const itemCountResponse = await pnp.sp
      //   .web
      //   .lists
      //   .getByTitle(libraryName)
      //   .select("ItemCount")
      //   .get();
      // const numberItems: number = itemCountResponse.ItemCount;
      // console.log(`The list ${libraryName} has ${numberItems} items.`);

      // get all files from root folder (1)
      // const allFilesRootFolder1: any[] = await pnp.sp
      //   .web
      //   .getFolderByServerRelativeUrl("Shared%20Documents")
      //   .files
      //   .get();
      // console.log(allFilesRootFolder1);

      // get all files from root folder (2)
      const allFilesRootFolder: any[] = await pnp.sp
        .web
        .lists
        .getByTitle(libraryName)
        .rootFolder
        .files
        .get();
      console.log(allFilesRootFolder);
      allFilesRootFolder.forEach((item) => {
        const size: number = item.Length;
        const sizeKB: number = size / 1024;
        docsInTreeMap = [...docsInTreeMap, { name: item.Name, value: sizeKB }];
        // tODO: Include URL clickable link into react-d3-treemap
        // ServerRelativeUrl
      });

      const allFolders: MyFolder[] = await pnp.sp
        .web
        .lists
        .getByTitle(libraryName)
        .rootFolder
        .folders
        .as(MyFolderCollection)
        // Filter forms and other empty folders
        .filter("ItemCount gt 0")
        .get(new SelectDecoratorsArrayParser<MyFolder>(MyFolder, true));


      for (let index: number = 0; index < allFolders.length; index++) {
        const folder: MyFolder = allFolders[index];
        // Todo. Work in bache
        const files = await pnp.sp
          .web
          .getFolderByServerRelativeUrl(folder.FolderUrl)
          .files
          .get();
        let folderFilesInTreeMap: ITreeMapNode[] = [];
        files.forEach((item) => {
          const size: number = item.Length;
          const sizeKB: number = size / 1024;
          folderFilesInTreeMap = [...folderFilesInTreeMap, { name: item.Name, value: sizeKB }];
          // tODO: Include URL clickable link into react-d3-treemap
          // ServerRelativeUrl
        });
        const folderNode: ITreeMapNode = { name: folder.FolderName, children: folderFilesInTreeMap };
        docsInTreeMap = [...docsInTreeMap, folderNode];
      }

      const data: ITreeMapNode = {
        "name": libraryName,
        "children": docsInTreeMap
      };

      // set our Component´s State
      this.setState({ ...this.state, data });
    } catch (error) {
      // set a new state conserving the previous state + the new error
      console.error(error);
      this.setState({
        ...this.state,
        errors: [...this.state.errors, "Error " + libraryName + ". Error: " + error]
      });
    }
  }

}
