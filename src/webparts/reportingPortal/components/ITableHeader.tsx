import * as React from 'react';
import { ITableHeaderProps } from './ITableHeaderProps';
import { ITableHeaderState } from './ITableHeaderState';
import { Row, Container, Col, Button } from 'react-bootstrap';
import 'bootstrap/dist/css/bootstrap.min.css';
import { SearchBox, ISearchBoxStyles } from 'office-ui-fabric-react/lib/SearchBox';
import styles from './ReportingPortal.module.scss';
const searchBoxStyles: Partial<ISearchBoxStyles> = { root: { width: 200 }, field: { border: "0 !important" } };
import { FontIcon } from '@fluentui/react/lib/Icon';

export default class ITableHeader extends React.Component<ITableHeaderProps, ITableHeaderState> {
    public constructor(props: ITableHeaderProps) {
        super(props);
    }
    public render(): React.ReactElement<ITableHeaderProps> {
        return (
            <React.Fragment>
                    <div className={"text-left col-sm-12 col-md-12 col-lg-3 col-xl-7 "+ styles.countDisplayMsg} dangerouslySetInnerHTML={this.props.totalCountDisplayMsg?{__html:  this.props.totalCountDisplayMsg }:null} />
                    {this.props.SelectedRecords && this.props.SelectedRecords.length > 0 ?
                        <Button color="primary" size="sm" className={"col-sm-12 col-md-12 col-lg-3 col-xl-2  m-2 " + styles.downloadSelectedBtn}
                            disabled={this.props.SelectedRecords && this.props.SelectedRecords.length > 0 ? false : true}
                            onClick={(e) => { this.props.DownloadSelected(e); }}
                        >
                            Download Selected
                </Button> :
                        <Button color="primary" size="sm" className={"col-sm-12 col-md-12 col-lg-3 col-xl-2  m-2 " + styles.downloadSelectedBtn}
                            onClick={(e) => {
                                e.preventDefault();
                                this.props.DownloadAll(e);
                            }}>Download All
                        </Button>
                    }
                    <SearchBox
                        value={this.props.SearchBoxInputText}
                        styles={searchBoxStyles}
                        placeholder="Search"
                        onChange={this.props.SearchTextChanged}
                        className="searchTxtBoxDiv col-sm-12 col-md-12 col-lg-2 col-xl-3 m-1 pl-0"
                    />
                
            </React.Fragment>
        );
    }
}
