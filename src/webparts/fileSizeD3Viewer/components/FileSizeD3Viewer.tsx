import * as React from 'react';

import styles from './FileSizeD3Viewer.module.scss';

// import PnP JS Core
import pnp from "sp-pnp-js";

// import models
import { MyDocument } from "../model/MyDocument";
import { MyDocumentCollection } from "../model/MyDocumentCollection";

// import custom parsers
import { SelectDecoratorsParser, SelectDecoratorsArrayParser } from "../parser/SelectDecoratorsParsers";


import { data } from "../data/mockData";
import TreeMap from "react-d3-treemap";
// Include its styles in you build process as well
import "react-d3-treemap/dist/react.d3.treemap.css";
import ContainerDimensions from 'react-container-dimensions';

// import React props and state
import { IFileSizeD3ViewerProps } from './IFileSizeD3ViewerProps';
import { IFileSizeD3ViewerState } from './IFileSizeD3ViewerState';

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
            valueUnit={"MB"}
          />
        }
      </ContainerDimensions>
      :
      <div>Loading...</div>
  }

  private async _readAllFilesSize(libraryName: string): Promise<void> {
    try {
      // query Item Count for the Library
      const docs: MyDocument[] = await pnp.sp
        .web
        .lists
        .getByTitle(libraryName)
        .items
        .select("FileLeafRef")
        //.as(MyDocumentCollection)
        //.get(new SelectDecoratorsArrayParser<MyDocument>(MyDocument));
        .get();

      //https://jquinto.sharepoint.com/sites/dev/_api/web/Lists/GetByTitle('Documents')/Items(1)
      debugger;
      const values = docs.map((item: MyDocument) => {
        const size: number = item.Size;
        const sizeKB: number = size / 1024;
        const name: string = item.Name;
        const id: string = item.Name;
        return { name, id, value: sizeKB };
      });
      const data = {
        "name": libraryName,
        "children": values
      };

      // Set our Component´s State
      this.setState({ ...this.state, data });
    } catch (error) {
      // set a new state conserving the previous state + the new error
      console.error(error);
      this.setState({
        ...this.state,
        errors: [...this.state.errors, "Error getting ItemCount for " + libraryName + ". Error: " + error]
      });
    }
  }

}
