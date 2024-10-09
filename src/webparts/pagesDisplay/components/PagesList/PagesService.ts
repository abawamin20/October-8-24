import { spfi, SPFI, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/views";
import "@pnp/sp/search";
import "@pnp/sp/fields";
import "@pnp/sp/site-users/web";

import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

import { WebPartContext } from "@microsoft/sp-webpart-base";

import { getColumnMaxWidth, getColumnMinWidth } from "../utils/columnUtils";
import { ConstructedFilter } from "./PanelComponent";
import { CellRender } from "../common/ColumnDetails";
export interface ITerm {
  Id: string;
  Name: string;
  parentId: string;
  Children?: ITerm[];
}

export interface TermSet {
  setId: string;
  terms: ITerm[];
}

export interface FilterDetail {
  filterColumn: string;
  filterColumnType: string;
  values: string[];
}

export interface IColumnInfo {
  InternalName: string;
  DisplayName: string;
  MinWidth: number;
  ColumnType: string;
  MaxWidth: number;
  OnRender?: (items: any) => JSX.Element;
}
class PagesService {
  private _sp: SPFI;

  constructor(private context: WebPartContext) {
    this._sp = spfi().using(SPFx(this.context));
  }

  /**
   * Fetch distinct values for a given column from a list of items.
   * @param {string} columnName - The name of the column to fetch distinct values for.
   * @param {any[]} values - The list of items to extract distinct values from.
   * @returns {Promise<string[] | ConstructedFilter[]>} - A promise that resolves to an array of distinct values.
   */
  getDistinctValues = async (
    columnName: string,
    columnType: string,
    values: any
  ): Promise<(string | ConstructedFilter)[]> => {
    try {
      const items = values; // The list of items to fetch distinct values from.

      // Extract distinct values from the column
      const distinctValues: (string | ConstructedFilter)[] = [];
      const seenValues = new Set<string | ConstructedFilter>(); // A set to keep track of seen values to avoid duplicates.

      items.forEach((item: any) => {
        switch (columnType) {
          case "TaxonomyFieldTypeMulti":
            if (item[columnName] && item[columnName].length > 0) {
              // Extract distinct values from the column
              item[columnName].forEach((category: any) => {
                const uniqueValue = category.Label;
                if (!seenValues.has(uniqueValue)) {
                  seenValues.add(uniqueValue);
                  distinctValues.push(uniqueValue);
                }
              });
            }
            break;
          case "DateTime":
            let uniqueDateValue = item[columnName]; // The value of the column for the current item.
            // Handle ISO date strings by extracting only the date part
            uniqueDateValue = new Date(uniqueDateValue)
              .toISOString()
              .split("T")[0];

            if (!seenValues.has(uniqueDateValue)) {
              seenValues.add(uniqueDateValue);
              distinctValues.push(uniqueDateValue);
            }
            break;
          case "User":
            const userValue = item[columnName];
            if (
              userValue &&
              userValue.Title &&
              !seenValues.has(userValue.Title)
            ) {
              seenValues.add(userValue.Title);
              const user: ConstructedFilter = {
                text: userValue.Title,
                value: userValue.Id,
              };
              distinctValues.push(user);
            }
            break;
          case "Number":
            const uniqueNumberValue = item[columnName]; // The value of the column for the current item.

            if (!seenValues.has(uniqueNumberValue)) {
              seenValues.add(uniqueNumberValue);
              distinctValues.push(uniqueNumberValue);
            }
            break;
          case "Choice":
            const uniqueChoiceValue = item[columnName]; // The value of the column for the current item.
            if (uniqueChoiceValue) {
              if (!seenValues.has(uniqueChoiceValue)) {
                seenValues.add(uniqueChoiceValue);
                distinctValues.push(uniqueChoiceValue);
              }
            }
            break;
          case "URL":
            const uniqueUrlChoiceValue = item[columnName]; // The value of the column for the current item.
            if (uniqueUrlChoiceValue && uniqueUrlChoiceValue.Url) {
              if (!seenValues.has(uniqueUrlChoiceValue.Url)) {
                seenValues.add(uniqueUrlChoiceValue.Url);
                distinctValues.push(uniqueUrlChoiceValue.Url);
              }
            }
            break;
          case "Computed":
            const uniqueCompChoiceValue = item[columnName]; // The value of the column for the current item.
            if (uniqueCompChoiceValue) {
              if (!seenValues.has(uniqueCompChoiceValue.split(".")[0])) {
                seenValues.add(uniqueCompChoiceValue.split(".")[0]);
                distinctValues.push(uniqueCompChoiceValue.split(".")[0]);
              }
            }
            break;
          default:
            const uniqueValue = item[columnName]; // The value of the column for the current item.
            if (uniqueValue) {
              if (!seenValues.has(uniqueValue)) {
                seenValues.add(uniqueValue);
                distinctValues.push(uniqueValue);
              }
            }
            break;
        }
      });

      return distinctValues;
    } catch (error) {
      console.error("Error fetching distinct values:", error);
      throw error;
    }
  };

  getDistinctValues2 = async (
    listId: string,
    columnName: string
  ): Promise<any[]> => {
    try {
      // Build the RenderListFilterData endpoint URL
      const filterDataEndpoint = `${this.context.pageContext.web.absoluteUrl}/_layouts/15/RenderListFilterData.aspx?FieldInternalName=${columnName}&ListId=${listId}`;

      const response: SPHttpClientResponse =
        await this.context.spHttpClient.get(
          filterDataEndpoint,
          SPHttpClient.configurations.v1
        );
      const responseData = await response.json();

      // Extract distinct values
      const distinctValues = responseData.filterData.map((item: any) => ({
        text: item.Value, // Display value
        value: item.Key, // Internal value
      }));

      return distinctValues;
    } catch (error) {
      console.error("Error fetching distinct values:", error);
      throw error;
    }
  };

  /**
   * Retrieves a page of filtered Site Pages items.
   *
   * @param orderBy The column to sort the items by. Defaults to "ModifiedDate".
   * @param isAscending Whether to sort in ascending or descending order. Defaults to true.
   * @param category The current category that is selected by the user
   * @param searchText Text to search for in the Title, Article ID, or Modified columns.
   * @param filters An array of FilterDetail objects to apply to the query.
   * @param columnInfos The columns information to include in the query.
   * @param pagesCache Cache to store fetched pages.
   * @param currentPageIndex The current page index.
   * @returns A promise that resolves with an array of items.
   */
  getFilteredPages = async (
    orderBy: string = "ModifiedDate",
    isAscending: boolean = true,
    category: string = "",
    searchText: string = "",
    filters: FilterDetail[],
    columnInfos: IColumnInfo[],
    pagesSize: any, // Fetch items per scroll
    nextPageUrl: string | null, // New parameter for the next page URL
    dateRangesIndex: number
  ) => {
    try {
      let queryUrl: string;

      // Check if nextPageUrl is provided (for pagination)
      if (nextPageUrl) {
        queryUrl = nextPageUrl;
      } else {
        const listUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('Site Pages')/items`;

        // Default columns to always include
        const allFieldsSet = new Set<string>();
        const expandFieldsSet = new Set<string>();

        // Populate column fields and expand fields from columnInfos
        columnInfos.forEach((col) => {
          if (col.ColumnType === "User") {
            allFieldsSet.add(`${col.InternalName}/Id`);
            allFieldsSet.add(`${col.InternalName}/Title`);
            expandFieldsSet.add(col.InternalName.split("/")[0]);
          } else if (
            col.ColumnType === "TaxonomyFieldTypeMulti" ||
            col.ColumnType === "TaxonomyFieldType"
          ) {
            allFieldsSet.add(col.InternalName);
          } else {
            allFieldsSet.add(`${col.InternalName}`);
          }
        });

        // Required fields
        const allFields: string[] = [
          "FileRef",
          "FileDirRef",
          "FSObjType",
          "Title",
          "Id",
          "FileLeafRef",
          "Article_x0020_ID",
          "Author/Title",
          "Editor/Title",
          "KnowledgeBaseLabel",
          "Modified",
          "ModifiedDate",
          "ModifiedBy",
          "Created",
          "CreatedDate",
          "CreatedBy",
        ];
        allFieldsSet.forEach((col) => allFields.push(col));

        const expandFields: string[] = ["Author", "Editor"];
        expandFieldsSet.forEach((field) => expandFields.push(field));

        // Build filter query
        let filterQuery = `KnowledgeBaseLabel eq '${category}' and FSObjType eq 0`;
        if (searchText) {
          filterQuery += ` and (substringof('${searchText}', Title) or Article_x0020_ID eq '${searchText}' or substringof('${searchText}', Modified))`;
        }

        // Date Range Filter
        if (dateRangesIndex && dateRangesIndex > 0) {
          const today = new Date();
          const currentYear = today.getFullYear();

          // Set the start date to the start of the year corresponding to (currentYear - dateRangesIndex)
          const startDate = new Date(
            Date.UTC(currentYear - dateRangesIndex, 0, 1)
          ).toISOString(); // January 1st of the target year

          // Set the end date to one year after the start date (start of the next year)
          const endDate = new Date(
            Date.UTC(currentYear - dateRangesIndex + 1, 0, 1)
          ).toISOString(); // January 1st of the next year
          filterQuery = `Modified ge datetime'${startDate}' and Modified le datetime'${endDate}' and ${filterQuery}`;
        }

        // Additional filters
        filters.forEach((filter) => {
          if (filter.values.length > 0) {
            switch (filter.filterColumnType) {
              case "DateTime":
                const dateFilters = filter.values
                  .map((value) => {
                    const startDate = new Date(value);
                    const endDate = new Date(value);
                    endDate.setDate(endDate.getDate() + 1);
                    return `${
                      filter.filterColumn
                    } ge datetime'${startDate.toISOString()}' and ${
                      filter.filterColumn
                    } lt datetime'${endDate.toISOString()}'`;
                  })
                  .join(" or ");
                if (dateFilters) filterQuery += ` and (${dateFilters})`;
                break;

              case "User":
                const userFilters = filter.values
                  .map((value) => `${filter.filterColumn}/Id eq '${value}'`)
                  .join(" or ");
                if (userFilters) filterQuery += ` and (${userFilters})`;
                break;

              case "URL":
                const urlFilters = filter.values
                  .map((value) => `${filter.filterColumn}/Url eq '${value}'`)
                  .join(" or ");
                if (urlFilters) filterQuery += ` and (${urlFilters})`;
                break;

              default:
                const columnFilters = filter.values
                  .map((value) => `${filter.filterColumn} eq '${value}'`)
                  .join(" or ");
                if (columnFilters) filterQuery += ` and (${columnFilters})`;
                break;
            }
          }
        });

        // Construct the final query URL with 'Paged=true'
        queryUrl = `${listUrl}?$select=${allFields.join(
          ","
        )}&$expand=${expandFields.join(
          ","
        )}&$filter=${filterQuery}&$top=${pagesSize}&$orderby=${orderBy} ${
          isAscending ? "asc" : "desc"
        }&$skiptoken=Paged=true`;
      }

      // Perform the API request
      const response: SPHttpClientResponse =
        await this.context.spHttpClient.get(
          queryUrl,
          SPHttpClient.configurations.v1
        );
      const jsonResponse = await response.json();

      if (jsonResponse.hasOwnProperty("error")) {
        return { error: jsonResponse.error.message };
      } else {
        // Use @odata.nextLink for pagination
        const nextPageLink = jsonResponse["@odata.nextLink"] || null;

        return {
          pages: jsonResponse.value,
          fetchednextPageUrl: nextPageLink, // Use next page URL for further pagination
        };
      }
    } catch (error) {
      console.error("Error fetching filtered pages", error);
      return { error: "error" };
    }
  };

  getFilteredPages2 = async (
    orderBy: string = "ModifiedDate",
    isDescending: boolean = false,
    category: string = "",
    searchText: string = "",
    filters: FilterDetail[],
    columnInfos: IColumnInfo[],
    pagesSize: any, // Fetch items per scroll
    lastPosition: number | null = 1
  ) => {
    try {
      let listUrl = `${
        this.context.pageContext.web.absoluteUrl
      }/_api/web/GetListUsingPath(DecodedUrl=@a1)/RenderListDataAsStream?@a1=%27%2Fsites%2FDevelopment%2FSitePages%27&TryNewExperienceSingle=TRUE&SortField=${orderBy}&SortDir=${
        isDescending ? "Desc" : "Asc"
      }`;

      // Default fields to include
      const allFieldsSet = new Set<string>();

      // Populate column fields and expand fields from columnInfos
      columnInfos.forEach((col) => {
        allFieldsSet.add(col.InternalName);
      });

      // Required fields
      const allFields: string[] = [
        "FileRef",
        "FileDirRef",
        "FSObjType",
        "Title",
        "Id",
        "FileLeafRef",
        "Article_x0020_ID",
        "Author/Title",
        "Editor/Title",
        "KnowledgeBaseLabel",
        "Modified",
        "ModifiedDate",
        "ModifiedBy",
        "Created",
        "CreatedDate",
        "CreatedBy",
      ];
      allFieldsSet.forEach((col) => allFields.push(col));
      let viewFieldsXML = "<ViewFields>";
      allFieldsSet.forEach((field) => {
        viewFieldsXML += `<FieldRef Name='${field}' />`;
      });
      viewFieldsXML += "</ViewFields>";

      // Build the CAML query structure
      let camlQuery = `<Where><And><Eq><FieldRef Name='FSObjType'/><Value Type='Number'>0</Value></Eq>`;

      // Filter by KnowledgeBaseLabel (category)
      if (category) {
        camlQuery += `<Eq><FieldRef Name='KnowledgeBaseLabel'/><Value Type='Text'>${category}</Value></Eq>`;
      }

      // Search text filtering
      if (searchText) {
        camlQuery += `<Or><Contains><FieldRef Name='Title'/><Value Type='Text'>${searchText}</Value></Contains>
                  <Eq><FieldRef Name='Article_x0020_ID'/><Value Type='Text'>${searchText}</Value></Eq></Or>`;
      }

      // Additional filters
      const buildFilterString = (filters: any[]) => {
        const filterParts: any[] = [];

        filters.forEach((filter, index) => {
          if (filter.values.length > 0) {
            // Iterate through the values of each filter
            filter.values.forEach((value: any) => {
              // Create the filter field, value, and type strings
              const filterField = `FilterField${index + 1}`;
              const filterValue = `FilterValue${index + 1}`;
              const filterType = `FilterType${index + 1}`;

              filterParts.push(
                `${filterField}=${encodeURIComponent(filter.filterColumn)}`
              );
              filterParts.push(`${filterValue}=${encodeURIComponent(value)}`);
              filterParts.push(
                `${filterType}=${encodeURIComponent(filter.filterColumnType)}`
              );
            });
          }
        });

        // Join the parts with '&' to create the final string
        return filterParts.join("&");
      };

      const filterString = buildFilterString(filters);

      camlQuery += `</And></Where>`;

      // Construct payload for RenderListDataAsStream
      let payload: {
        parameters: any;
      } = {
        parameters: {
          AllowMultipleValueFilterForTaxonomyFields: true,
          AddRequiredFields: false,
          RequireFolderColoringFields: true,
          ViewXml: `<View Scope="RecursiveAll">${viewFieldsXML}<Query>${camlQuery}</Query><RowLimit Paged="TRUE">${pagesSize}</RowLimit></View>`,
        },
      };
      if (lastPosition) {
        payload.parameters.Paging = `Paged=TRUE&p_ID=${lastPosition}`;
      }
      // Perform the API request
      const response = await this.context.spHttpClient.post(
        filterString ? `${listUrl}&${filterString}` : listUrl,
        SPHttpClient.configurations.v1,
        {
          body: JSON.stringify(payload),
        }
      );
      const jsonResponse = await response.json();

      if (jsonResponse.hasOwnProperty("error")) {
        return { error: jsonResponse.error.message };
      } else {
        const lastPosition = jsonResponse.Row
          ? parseInt(jsonResponse.Row[jsonResponse.LastRow - 1].ID)
          : null;

        return {
          pages: jsonResponse.Row,
          lastPosition: lastPosition, // Store this for the next request
        };
      }
    } catch (error) {
      console.error("Error fetching filtered pages", error);
      return { error: "error" };
    }
  };

  /**
   * Retrieves the columns for a specified view in the SharePoint list.
   */
  public async getColumns(viewId: string): Promise<IColumnInfo[]> {
    const fields = await this._sp.web.lists
      .getByTitle("Site Pages")
      .views.getById(viewId)
      .fields();

    // Fetching detailed field information to get both internal and display names
    const fieldDetailsPromises = fields.Items.map((field: any) =>
      this._sp.web.lists
        .getByTitle("Site Pages")
        .fields.getByInternalNameOrTitle(field)()
    );

    const fieldDetails = await Promise.all(fieldDetailsPromises);

    return fieldDetails.map((field: any) => ({
      InternalName: field.InternalName,
      DisplayName: field.Title,
      ColumnType: field.TypeAsString,
      MinWidth: getColumnMinWidth(field.InternalName),
      MaxWidth: getColumnMaxWidth(field.InternalName),
      OnRender: (item: any) =>
        CellRender({
          columnName: field.InternalName,
          columnType: field.TypeAsString,
          item,
          context: this.context,
        }),
    }));
  }

  /**
   * Retrieves the details of a SharePoint list by its name.
   * @param {string} listName - The name of the list to retrieve details for.
   * @returns {Promise<any>} - A promise that resolves to the list details.
   */
  public async getListDetailsByName(listName: string): Promise<any> {
    try {
      const list = await this._sp.web.lists.getByTitle(listName)();
      return list;
    } catch (error) {
      console.error(`Error retrieving list details for ${listName}:`, error);
      throw new Error(`Error retrieving list details for ${listName}`);
    }
  }

  public async createListItem(itemData: any, listTitle: string): Promise<any> {
    try {
      const addedItem = await this._sp.web.lists
        .getByTitle(listTitle) // Get list by title
        .items.add(itemData);
      return addedItem;
    } catch (error) {
      console.error("Error creating list item: ", error);
      throw error;
    }
  }

  async getByUrl(url: string): Promise<any> {
    try {
      return this.context.spHttpClient
        .get(url, SPHttpClient.configurations.v1)
        .then((response: SPHttpClientResponse) => {
          return response.json();
        });
    } catch (error) {}
  }

  public async getUniqueFilterOptions(
    listTitle: string,
    columnName: string
  ): Promise<any[]> {
    const apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listTitle}')/RenderListFilterData`;

    const body = JSON.stringify({
      // filterLink: `/Lists/${listTitle}/AllItems.aspx`,
      fieldInternalName: columnName,
    });

    try {
      const response = await this.context.spHttpClient.post(
        apiUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose",
          },
          body: body,
        }
      );

      const result = await response.json();
      const filterOptions: { text: string; value: string }[] =
        result.d.FilterData.Row.map((item: any) => ({
          text: item.DisplayValue, // Map DisplayValue to text
          value: item.Key, // Map Key to value
        }));

      return filterOptions;
    } catch (error) {
      console.error("Error fetching filter options:", error);
      return [];
    }
  }
}

export default PagesService;
