import * as React from "react";
import {
  DetailsList,
  // DetailsListLayoutMode,
  SelectionMode,
  IColumn,
  DetailsHeader,
  Selection,
  IDetailsListStyles,
  ConstrainMode,
} from "@fluentui/react/lib/DetailsList";
import { mergeStyles, mergeStyleSets } from "@fluentui/react";
import "./styles.css";
import { IColumnInfo } from "../PagesList/PagesService";
import { WebPartContext } from "@microsoft/sp-webpart-base";

const gridStyles: Partial<IDetailsListStyles> = {
  root: {
    overflowX: "scroll",
    selectors: {
      "& [role=grid]": {
        display: "flex",
        flexDirection: "column",
        alignItems: "start",
        // height: "600px",
      },
    },
  },
  headerWrapper: {
    flex: "0 0 auto",
  },
  contentWrapper: {
    flex: "1 1 auto",
    overflow: "hidden",
  },
};

const classNames = mergeStyleSets({
  header: {
    margin: 0,
  },
  row: {
    flex: "0 0 auto",
  },
  focusZone: {
    height: "100%",
    maxHeight: "600px",
    overflowY: "scroll",
    overflowX: "hidden",
    "scrollbar-width": "thin",
    "scrollbar-color": "#f5f5f5",
  },
  selectionZone: {
    height: "100%",
    overflow: "hidden",
  },
});

const customHeaderClass = mergeStyles({
  backgroundColor: "#efefef",
  color: "white",
  paddingTop: 0,
  paddingBottom: 0,
  selectors: {
    "& .ms-DetailsHeader": {
      backgroundColor: "#0078d4",
      borderBottom: "1px solid #ccc",
    },
  },
});

export interface IReusableDetailListcomponentsProps {
  columns: (
    columns: IColumnInfo[],
    context: WebPartContext,
    currentUser: any,
    onColumnClick: any,
    sortBy: string,
    isDecending: boolean,
    setShowFilter: (column: IColumn, columnType: string) => void
  ) => IColumn[];
  columnInfos: IColumnInfo[];
  currentUser: any;
  context: WebPartContext;
  setShowFilter: (column: IColumn, columnType: string) => void;
  updateSelection: (selection: Selection) => void;
  items: any[];
  sortPages: (column: IColumn, isAscending: boolean) => void;
  sortBy: string;
  siteUrl: string;
  isDecending: boolean;
  loadMoreItems: () => void; // New prop to load more items
  initialScrollTop: number;
  updateScrollPosition: (scrollTop: number) => void;
}

export interface IReusableDetailListcomponentsState {}
export class ReusableDetailList extends React.Component<
  IReusableDetailListcomponentsProps,
  IReusableDetailListcomponentsState
> {
  private _selection: Selection;
  private containerRef: React.RefObject<HTMLDivElement>;
  private _selectionChanged: boolean = false;
  constructor(props: IReusableDetailListcomponentsProps) {
    super(props);
    this._selection = new Selection({
      onSelectionChanged: () => {
        this._selectionChanged = true;
        this.props.updateSelection(this._selection);
      },
      getKey: this._getKey,
    });

    this.state = {
      isLoading: false,
    };

    this.containerRef = React.createRef(); // Ref for the scrollable container
  }

  componentDidMount(): void {
    const { initialScrollTop } = this.props;

    const focusZoneElement = document.querySelectorAll(".focusCustomClass");

    if (this._selectionChanged) {
      this._selectionChanged = false;
      return;
    }
    if (focusZoneElement.length > 0) {
      // Restore scroll position if props change
      focusZoneElement[0].scrollTop = initialScrollTop;
    }
    this.setState({ isLoading: true });
    setTimeout(() => {
      this.setState({ isLoading: false });
    }, 100);
  }

  componentDidUpdate(prevProps: any, prevState: any) {
    const { initialScrollTop } = this.props;
    const focusZoneElement = document.querySelectorAll(".focusCustomClass");

    if (this._selectionChanged) {
      this._selectionChanged = false;
      return;
    }
    if (focusZoneElement.length > 0) {
      // Restore scroll position if props change
      focusZoneElement[0].scrollTop = initialScrollTop;
    }
  }

  handleScroll = (event: React.UIEvent<HTMLElement>) => {
    const container = event.currentTarget; // The scrollable element itself
    const scrollTop = container.scrollTop;
    const scrollHeight = container.scrollHeight;
    const clientHeight = container.clientHeight;

    // Only load more items if scrolled close to the bottom
    if (scrollTop + clientHeight + 5 >= scrollHeight) {
      this.props.updateScrollPosition(scrollTop);
      this.props.loadMoreItems();
    }
  };
  public render() {
    const {
      columnInfos,
      currentUser,
      context,
      columns,
      items,
      sortPages,
      sortBy,
      isDecending,
      setShowFilter,
    } = this.props; // Custom selection logic

    const focusZoneProps = {
      className: `${classNames.focusZone} focusCustomClass`,
      "data-is-scrollable": "true",
      onScroll: this.handleScroll,
    } as React.HTMLAttributes<HTMLElement>;

    return (
      <div ref={this.containerRef}>
        <DetailsList
          compact={true}
          items={items}
          columns={columns(
            columnInfos,
            context,
            currentUser,
            sortPages,
            sortBy,
            isDecending,
            setShowFilter
          )}
          selectionMode={SelectionMode.single}
          setKey="set"
          selection={this._selection}
          constrainMode={ConstrainMode.unconstrained}
          onRenderDetailsHeader={this._onRenderDetailsHeader}
          selectionPreservedOnEmptyClick
          styles={gridStyles}
          ariaLabelForSelectionColumn="Toggle selection"
          ariaLabelForSelectAllCheckbox="Toggle selection for all items"
          checkButtonAriaLabel="select row"
          onItemInvoked={this._onItemInvoked}
          focusZoneProps={focusZoneProps}
          selectionZoneProps={{
            className: classNames.selectionZone,
          }}
        />
      </div>
    );
  }

  private _getKey(item: any, index?: number): string {
    return item.key || index?.toString() || "";
  }

  private _onItemInvoked = (item: any): void => {
    window.open(`${this.props.siteUrl}${item.FileRef}`, "_blank");
  };

  private _onRenderDetailsHeader = (props: any) => {
    if (!props) {
      return null;
    }

    return <DetailsHeader {...props} styles={{ root: customHeaderClass }} />;
  };
}
