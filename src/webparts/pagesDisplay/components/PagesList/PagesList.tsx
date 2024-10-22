import * as React from "react";
import { Dialog, DialogFooter, Spinner } from "@fluentui/react";
import { SPHttpClient } from "@microsoft/sp-http";
import { ReusableDetailList } from "../common/ReusableDetailList";
import PagesService, { FilterDetail, IColumnInfo } from "./PagesService";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { PagesColumns, subscriptionCache } from "./PagesColumns";
import { DefaultButton, IColumn, Icon, Selection } from "@fluentui/react";
import { makeStyles, useId, Input } from "@fluentui/react-components";
import styles from "./pages.module.scss";
import "./pages.css";
import { FilterPanelComponent } from "./PanelComponent";
import ListForm from "../Forms/ListForm";

interface SuccessResponse {
  pages: any;
  nextPageLink: string | null;
}

interface ErrorResponse {
  error: string;
}

export interface IPagesListProps {
  context: WebPartContext;
  selectedViewId: string;
  feedbackPageUrl: string;
}

const useStyles = makeStyles({
  root: {
    display: "flex",
    gap: "2px",
    maxWidth: "400px",
    alignItems: "center",
  },
});

const PagesList = (props: IPagesListProps) => {
  const subscribeLink: string = "/_layouts/15/SubNew.aspx";
  const alertLink: string = "/_layouts/15/mySubs.aspx";

  // Destructure the props
  const { context, selectedViewId } = props;

  /**
   * State variables for the component.
   */

  // Options for the page size dropdown
  const [pageSizeOption] = React.useState<number[]>([
    10, 20, 40, 60, 80, 100, 200, 300, 400, 500,
  ]);

  const [hideFeedBackDialog, setHideFeedBackDialog] = React.useState(true);

  const toggleHideFeedbackDialog = () => {
    setHideFeedBackDialog(!hideFeedBackDialog);
  };
  const [hideAlertMeDialog, setHideAlertMeDialog] = React.useState(true);
  const [hideManageAlertDialog, setHideManageAlertDialog] =
    React.useState(true);

  const toggleHideAlertMeDialog = () => {
    setHideAlertMeDialog(!hideAlertMeDialog);
  };
  const toggleHideManageAlertDialog = () => {
    setHideManageAlertDialog(!hideManageAlertDialog);
  };

  const [columnInfos, setColumnInfos] = React.useState<IColumnInfo[]>([]);

  // The search text for filtering pages
  const [searchText, setSearchText] = React.useState<string>(""); // Initially set to empty string

  // The list of pages
  const [pages, setPages] = React.useState<any[]>([]); // Initially set to empty array

  // The selected category
  const [catagory, setCatagory] = React.useState<string | null>(null); // Initially set to empty string
  const [isLoading, setIsLoading] = React.useState<boolean>(false); // Initially set to empty string

  // The column to sort by
  const [sortBy, setSortBy] = React.useState<string>(""); // Initially set to empty string
  const [scrollTop, setScrollTop] = React.useState<number>(0); // Initially set to empty string

  const [hasNextPage, setHasNextPage] = React.useState<boolean>(false);
  const [nextPageLink, setNextPageLink] = React.useState<string | null>(null);

  // The number of items to display per page
  const [pageSize, setPageSize] = React.useState<number>(500); // Initially set to 20

  // The total number of items
  const [totalItems, setTotalItems] = React.useState<number>(0); // Initially set to 0

  // The sorting order
  const [isDescending, setIsDescending] = React.useState<boolean>(true); // Initially set to false

  // Whether to show the filter panel
  const [showFilter, setShowFilter] = React.useState<boolean>(false); // Initially set to false

  // The column to filter by
  const [filterColumn, setFilterColumn] = React.useState<string>(""); // Initially set to empty string

  // The type of column to filter by
  const [filterColumnType, setFilterColumnType] = React.useState<string>(""); // Initially set to empty string

  // The filter details
  const [filterDetails, setFilterDetails] = React.useState<FilterDetail[]>([]); // Initially set to empty array

  // The filter details
  const [selectionDetails, setSelectionDetails] = React.useState<any | []>([]);
  // The filter details
  const [listId, setListId] = React.useState<string>("");
  const [currentUser, setCurrentUser] = React.useState<any>(null);
  const [viewId, setViewId] = React.useState<string>("");

  // Create an instance of the PagesService class
  const pagesService = new PagesService(context);

  // Get a unique id for the input field
  const inputId = useId("input");

  // Get the styles for the input field
  const inputStyles = useStyles();

  const subscribeIframeRef = React.useRef<HTMLIFrameElement>(null);

  const updateSubscriptionOfSelectedPage = async () => {
    try {
      const selectedItem = selectionDetails[0];
      const itemKey = `SitePages_${selectedItem.ID}`;
      const pageTitle = `Site Pages: ${selectedItem.FileLeafRef.replace(
        /'/g,
        "''"
      )}}`; // Ensure correct page title format
      const alertResponse = await context.spHttpClient.get(
        `${context.pageContext.web.absoluteUrl}/_api/web/alerts?$filter=UserId eq ${currentUser.Id} and Title eq '${pageTitle}'`,
        SPHttpClient.configurations.v1
      );
      const alertData = await alertResponse.json();
      console.log(alertData);
      console.log(itemKey);
      // Set subscribed based on the presence of alerts
      const isSubscribed = alertData.value.length > 0;

      // Update the cache with the fetched status
      subscriptionCache.set(itemKey, isSubscribed);
    } catch (error) {}
  };
  React.useEffect(() => {
    const checkIframeUrl = () => {
      const iframe = subscribeIframeRef.current;
      if (iframe && iframe.contentWindow) {
        const currentUrl = iframe.contentWindow.location.href;

        if (
          currentUrl.indexOf("blank") === -1 &&
          currentUrl.indexOf(subscribeLink) === -1
        ) {
          setHideAlertMeDialog(true);
          updateSubscriptionOfSelectedPage();
          setSelectionDetails([]);
          setPages([]);
          setTotalItems(0);
          setScrollTop(0);
          setHasNextPage(true);
          setNextPageLink(null);

          fetchPages(
            pageSize,
            "ModifiedDate",
            true,
            searchText,
            catagory,
            filterDetails,
            columnInfos,
            [],
            null
          );
        }
      }
    };

    // Check the URL every 2 seconds
    const intervalId = setInterval(checkIframeUrl, 2000);

    // Clean up the interval on component unmount
    return () => clearInterval(intervalId);
  }, [
    subscribeLink,
    setHideAlertMeDialog,
    catagory,
    catagory,
    selectionDetails,
  ]);

  /**
   * Resets the filters by clearing the checked items and
   * calling the applyFilters function with an empty filter detail.
   */
  const resetFilters = () => {
    // Clear the filter details
    setFilterDetails([]);
    setTotalItems(0);
    setScrollTop(0);
    setHasNextPage(true);
    setPages([]);
    setNextPageLink(null);

    // Clear the search text
    setSearchText("");

    // Call the fetchPages function with the default arguments
    fetchPages(
      pageSize,
      "ModifiedDate",
      true,
      "",
      catagory,
      [],
      columnInfos,
      [],
      null
    );
  };

  /**
   * Fetches the paginated pages based on the given parameters.
   *
   * @param {number} [pageSizeAmount=pageSize] - The number of items per page. Defaults to the `pageSize` state variable.
   * @param {string} [sortBy="ModifiedDate"] - The column to sort by. Defaults to "ModifiedDate".
   * @param {boolean} [isSortedDescending=isDescending] - Whether to sort in descending order. Defaults to the `isDescending` state variable.
   * @param {string} [searchText=""] - The search text to filter by. Defaults to an empty string.
   * @param {string} [category=catagory] - The category to filter by. Defaults to the `catagory` state variable.
   * @param {FilterDetail[]} filterDetails - The filter details to apply.
   *
   * @return {Promise<void>} - A promise that resolves when the paginated pages are fetched.
   */
  const fetchPages = async (
    pageSizeAmount: number = pageSize, // Always fetch 50 items per request
    sortBy: string = "ModifiedDate",
    isSortedDescending: boolean = isDescending,
    searchText: string = "",
    category: string | null = catagory,
    filterDetails: FilterDetail[] = [],
    columns: IColumnInfo[] = columnInfos,
    currentPages: any[] = pages,
    nextPageLinkFetched: string | null = nextPageLink
  ): Promise<any[]> => {
    // Set loading state and clear selection
    setIsLoading(true);
    setSelectionDetails([]);

    try {
      // Fetch pages with current lastPositionFetched or date range
      const res: SuccessResponse | ErrorResponse =
        await pagesService.getFilteredPages2(
          sortBy,
          isSortedDescending,
          category as string,
          searchText,
          filterDetails,
          columns,
          pageSizeAmount,
          nextPageLinkFetched
        );

      let { pages: fetchedPages, nextPageLink } = res as SuccessResponse;

      // Combine the new pages with the current pages
      const finalPages = [...currentPages, ...fetchedPages];
      setPages(finalPages);
      setTotalItems(finalPages.length);
      // Handle nextPageUrl for pagination
      if (nextPageLink) {
        setHasNextPage(true);
        setNextPageLink(nextPageLink);
      } else {
        // End of pagination if no nextPageUrl
        setHasNextPage(false);
        setNextPageLink(null);
      }

      // If the number of fetched pages is less than 50 and we're not in threshold error mode
      if (fetchedPages.length < pageSizeAmount) {
        // Only try to fetch more from date range if we're in threshold error mode

        setHasNextPage(false);
      }

      setIsLoading(false);
      return finalPages; // Return all accumulated pages
    } catch (error) {
      // Handle unexpected errors
      console.error("Error fetching pages:", error);
      setIsLoading(false);
      throw error; // Re-throw error for handling in the calling function
    }
  };

  /**
   * Applies the given filter details to filter the pages
   *
   * @param {FilterDetail} filterDetail - The filter detail object containing the filter details
   */
  const applyFilters = (filterDetail: FilterDetail): void => {
    /**
     * Updates the current filter details state with the new filter detail,
     * or removes the filter detail if the values array is empty.
     *
     */
    let currentFilters: FilterDetail[] = filterDetails;

    if (filterDetail.values.length === 0) {
      currentFilters = filterDetails.filter(
        (item) => item.filterColumn !== filterDetail.filterColumn
      );
    } else
      currentFilters = [
        ...filterDetails.filter(
          (item) => item.filterColumn !== filterDetail.filterColumn
        ),
        filterDetail,
      ];
    setNextPageLink(null);
    setFilterDetails(currentFilters);

    fetchPages(
      pageSize, // Page size
      "ModifiedDate", // Sorting criteria
      true, // Sorting order (ascending/descending)
      searchText, // Search text
      catagory, // Category (assuming this is another state or prop)
      currentFilters, // Updated filter details,
      columnInfos,
      [],
      null
    );
  };

  /**
   * Sort the pages list based on the specified column.
   *
   * @param {IColumn} column - The column to sort by.
   */
  const sortPages = (column: IColumn) => {
    // Set the sort by column state
    setSortBy(column.fieldName as string);

    // If the column is the same as the current sort by column, toggle the sort order
    if (column.fieldName === sortBy) {
      setIsDescending(!isDescending);
    } else {
      // Otherwise, set the sort order to descending
      setIsDescending(true);
    }

    // Fetch the pages list with the new sort criteria
    fetchPages(
      pageSize, // Page size
      column.fieldName, // Sorting criteria
      column.isSortedDescending, // Sorting order (ascending/descending)
      searchText, // Search text
      catagory, // Category (assuming this is another state or prop)
      filterDetails, // Filter details
      columnInfos,
      [],
      null
    );
  };

  /**
   * Handles the search functionality by fetching pages with specified parameters.
   */
  const handleSearch = () => {
    fetchPages(
      pageSize, // Page size
      "ModifiedDate", // Sorting criteria
      true, // Sorting order (ascending/descending)
      searchText, // Search text
      catagory, // Category
      filterDetails, // Filter details
      columnInfos,
      [],
      null
    );
  };

  /**
   * Handles the change event of the page size dropdown.
   *
   * This function is triggered when the user selects a new page size from the dropdown.
   * It updates the page size state and calls the `handlePageChange` function to update
   * the paginated data.
   *
   * @function handlePageSizeChange
   * @memberof PagesList
   *
   * @param {any} e - The event object.
   * @return {void}
   */
  const handlePageSizeChange = (e: any) => {
    // Update the page size state
    setPageSize(e.target.value);
    setPages([]);
    setTotalItems(0);
    setScrollTop(0);

    setHasNextPage(true);
    setNextPageLink(null);
    // Handle the page change with the new page size

    fetchPages(
      e.target.value,
      "ModifiedDate",
      true,
      searchText,
      catagory,
      filterDetails,
      columnInfos,
      [],
      null
    );
  };

  /**
   * Dismisses the filter panel.
   * Sets the showFilter state to false.
   *
   * @function dismissPanel
   * @memberof PagesList
   * @returns {void}
   */
  const dismissPanel = (): void => {
    setShowFilter(false);
  };

  const getColumns = async (selectedViewId: string) => {
    const columns = await pagesService.getColumns(selectedViewId);

    setColumnInfos(columns);

    return columns;
  };

  React.useEffect(() => {
    const handleEvent = (e: any) => {
      if (columnInfos.length > 0) {
        const selectedCategory = e.detail;

        if (selectedCategory && selectedCategory != "") {
          setCatagory(selectedCategory);

          fetchPages(
            pageSize,
            "ModifiedDate",
            true,
            searchText,
            selectedCategory,
            filterDetails,
            columnInfos,
            [],
            null
          );
          setSelectionDetails([]);
          setPageSize(pageSize);
        }
      }
    };

    pagesService.getListDetailsByName("Site Pages").then((res) => {
      setListId(res.Id);
    });

    window.addEventListener("catagorySelected", handleEvent);
  }, [columnInfos]);

  React.useEffect(() => {
    const fetchCurrentUser = async () => {
      try {
        const currentUserResponse = await context.spHttpClient.get(
          `${context.pageContext.web.absoluteUrl}/_api/web/currentuser`,
          SPHttpClient.configurations.v1
        );
        const userData = await currentUserResponse.json();
        setCurrentUser(userData);
      } catch (error) {
        console.error("Error fetching current user:", error);
      }
    };
    fetchCurrentUser();
  }, []);

  React.useEffect(() => {
    if (viewId !== selectedViewId) {
      setViewId(selectedViewId);
      getColumns(selectedViewId).then((col) => {
        if (catagory && catagory != "") {
          fetchPages(
            pageSize,
            "ModifiedDate",
            true,
            searchText,
            catagory,
            filterDetails,
            col,
            [],
            null
          );
        }
      });
    }
  }, [selectedViewId]);

  return (
    <div className="w-pageSize0 detail-display">
      {showFilter && (
        <FilterPanelComponent
          isOpen={showFilter}
          headerText="Filter Articles"
          applyFilters={applyFilters}
          dismissPanel={dismissPanel}
          selectedItems={
            [...filterDetails].filter(
              (item) => item.filterColumn === filterColumn
            )[0] || { filterColumn: "", values: [] }
          }
          columnName={filterColumn}
          columnType={filterColumnType}
          pagesService={pagesService}
          data={pages}
          listId={listId}
        />
      )}
      <div className={`${styles.top}`}>
        <div
          className={`${styles["first-section"]} d-flex justify-content-between align-items-end py-2 px-2`}
        >
          <span className={`fs-4 ${styles["knowledgeText"]}`}>
            {catagory && <span className="">{catagory}</span>}
          </span>
          <div className={`${inputStyles.root} d-flex align-items-center me-2`}>
            <Input
              id={inputId}
              value={searchText}
              onChange={(e) => setSearchText(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === "Enter") {
                  handleSearch();
                }
              }}
              placeholder="Search"
            />
          </div>
        </div>

        <div
          className={`d-flex justify-content-between align-items-center fs-5 px-2 my-2`}
        >
          <span>Articles /</span>
          {totalItems > 0 ? (
            <div className="d-flex align-items-center">
              {selectionDetails && selectionDetails.length > 0 && (
                <DefaultButton
                  className="me-2"
                  onClick={() => {
                    toggleHideAlertMeDialog();
                  }}
                >
                  <span className="d-flex align-items-center">
                    <Icon iconName="Ringer" className="me-2" />
                    Alert Me
                  </span>
                </DefaultButton>
              )}
              {selectionDetails && selectionDetails.length > 0 && (
                <DefaultButton
                  className="me-2"
                  onClick={() => {
                    toggleHideManageAlertDialog();
                  }}
                >
                  <span className="d-flex align-items-center">
                    <Icon iconName="EditNote" className="me-2" />
                    Manage My Alerts
                  </span>
                </DefaultButton>
              )}
              {selectionDetails && selectionDetails.length > 0 && (
                <DefaultButton
                  className="me-2"
                  onClick={() => {
                    toggleHideFeedbackDialog();
                  }}
                >
                  <span className="d-flex align-items-center">
                    <Icon iconName="Feedback" className="me-2" />
                    Add Feedback
                  </span>
                </DefaultButton>
              )}
              {filterDetails && filterDetails.length > 0 && (
                <DefaultButton
                  onClick={() => {
                    resetFilters();
                  }}
                >
                  Clear
                </DefaultButton>
              )}
              <span className="ms-2 fs-6">Results ({totalItems})</span>
            </div>
          ) : (
            <span className="fs-6">No articles to display</span>
          )}
        </div>
      </div>

      {isLoading ? (
        <div style={{ textAlign: "center", minHeight: "300px" }}>
          <Spinner label="Articles are being loaded..." />
        </div>
      ) : (
        <div>
          <ReusableDetailList
            items={pages}
            context={context}
            columns={PagesColumns}
            columnInfos={columnInfos}
            currentUser={currentUser}
            setShowFilter={(column: IColumn, columnType: string) => {
              setShowFilter(!showFilter);
              setFilterColumn(column.fieldName as string);
              setFilterColumnType(columnType);
            }}
            updateSelection={(selection: Selection) => {
              setSelectionDetails(selection.getSelection());
            }}
            sortPages={sortPages}
            sortBy={sortBy}
            siteUrl={window.location.origin}
            isDecending={isDescending}
            loadMoreItems={() => {
              hasNextPage &&
                fetchPages(
                  pageSize,
                  "ModifiedDate",
                  true,
                  searchText,
                  catagory,
                  filterDetails,
                  columnInfos,
                  pages,
                  nextPageLink
                );
            }}
            initialScrollTop={scrollTop}
            updateScrollPosition={(scrollTop: number) => {
              setScrollTop(scrollTop);
            }}
          />
        </div>
      )}
      <div className="d-flex justify-content-end">
        <div
          className="d-flex align-items-center my-1"
          style={{
            fontSize: "13px",
          }}
        >
          <div className="d-flex align-items-center me-3">
            <span className={`me-2 ${styles.blueText}`}>Items / Page </span>
            <select
              className="form-select"
              value={pageSize}
              onChange={handlePageSizeChange}
              name="pageSize"
              style={{
                width: 80,
                height: 35,
              }}
            >
              {pageSizeOption.map((pageSize) => {
                return (
                  <option key={pageSize} value={pageSize}>
                    {pageSize}
                  </option>
                );
              })}
            </select>
          </div>
        </div>
      </div>

      <Dialog
        hidden={hideFeedBackDialog}
        onDismiss={toggleHideFeedbackDialog}
        modalProps={{
          isBlocking: false,
        }}
        maxWidth="90vw"
        minWidth="60vw"
      >
        <ListForm
          articleId={
            selectionDetails[0] && selectionDetails[0].Article_x0020_ID
          }
          title={selectionDetails[0] && selectionDetails[0].Title}
          name={selectionDetails[0] && selectionDetails[0].FileLeafRef}
          link={
            selectionDetails[0] &&
            `${window.location.origin}${selectionDetails[0].FileRef}`
          }
          hideDialog={() => setHideFeedBackDialog(true)}
          pageService={pagesService}
          currentUser={currentUser}
          catagory={catagory}
          createdBy={selectionDetails[0] && selectionDetails[0].CreatedBy}
          modifiedBy={selectionDetails[0] && selectionDetails[0].ModifiedBy}
          createdDate={selectionDetails[0] && selectionDetails[0].Created}
          modifiedDate={selectionDetails[0] && selectionDetails[0].Modified}
        />
      </Dialog>
      <Dialog
        hidden={hideAlertMeDialog}
        onDismiss={toggleHideAlertMeDialog}
        modalProps={{
          isBlocking: false,
        }}
        maxWidth="90vw"
        minWidth="60vw"
      >
        <iframe
          ref={subscribeIframeRef}
          src={`${
            context.pageContext.web.absoluteUrl
          }${subscribeLink}?List=${listId}&Id=${
            selectionDetails[0] && selectionDetails[0].ID
          }`}
          width="100%"
          height="600px"
          style={{ border: "none" }}
        ></iframe>

        <DialogFooter>
          <DefaultButton
            onClick={() => {
              setHideAlertMeDialog(true);
            }}
            text="Close"
          />
        </DialogFooter>
      </Dialog>
      <Dialog
        hidden={hideManageAlertDialog}
        onDismiss={toggleHideManageAlertDialog}
        modalProps={{
          isBlocking: false,
        }}
        maxWidth="90vw"
        minWidth="60vw"
      >
        <iframe
          src={`${context.pageContext.web.absoluteUrl}${alertLink}`}
          width="100%"
          height="600px"
          style={{ border: "none" }}
          id="alertFrame"
        ></iframe>

        <DialogFooter>
          <DefaultButton
            onClick={() => {
              setHideManageAlertDialog(true);

              fetchPages(
                pageSize,
                "ModifiedDate",
                true,
                searchText,
                catagory,
                filterDetails,
                columnInfos,
                [],
                null
              );

              subscriptionCache.clear();
            }}
            text="Close"
          />
        </DialogFooter>
      </Dialog>
    </div>
  );
};

export default PagesList;
