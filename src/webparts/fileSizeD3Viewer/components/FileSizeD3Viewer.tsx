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
      items: [],
      errors: []
    };
  }

  public render(): React.ReactElement<IFileSizeD3ViewerProps> {
    return (
      <div>
        <ContainerDimensions>
          {({ width, height }) =>
            <TreeMap
              width={width}
              height={350}
              data={data}
              valueUnit={"MB"}
            />
          }
        </ContainerDimensions>
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

  private async _readAllFilesSize(libraryName: string): Promise<void> {
    try {
      // query Item Count for the Library
      const docs: MyDocument[] = await pnp.sp
        .web
        .lists
        .getByTitle(libraryName)
        .items
        .as(MyDocumentCollection)
        .get(new SelectDecoratorsArrayParser<MyDocument>(MyDocument));

        // Set our Component´s State
        this.setState({ ...this.state, items: docs });
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
