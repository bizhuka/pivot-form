<template>
  <v-container>

    <v-container class="d-flex justify-center" v-if="columns.length === 0">
      <v-alert type="info" dense>
        <v-row>
          <v-col cols="12">
            <p>There is no active worksheet with an autofilter.</p>
          </v-col>
        </v-row>
      </v-alert>
    </v-container>

    <v-container class="d-flex justify-center" v-if="columns.length > 0">
      <v-btn color="primary" @click="dialog = true" class="mb-4 mr-2">Save Filter</v-btn>
      <v-btn color="primary" @click="deleteFilter" class="mb-4 ml-2">Delete Filter</v-btn>
    </v-container>

    <v-select :items="filters" v-model="currentFilter" item-text="title" item-value="title" label="Select Filter"
      density="compact" @update:modelValue="onApplyChange" clearable v-if="columns.length > 0"></v-select>

    <!-- Dialog with obligatory input field -->
    <v-dialog v-model="dialog" max-width="200em">
      <v-card>
        <v-card-title>Save as</v-card-title>
        <v-card-text>
          <v-form ref="form" v-model="valid">
            <v-text-field v-model="inputValue" :rules="[v => !!v || 'Input is required']" label="Filter name"
              required></v-text-field>
          </v-form>
        </v-card-text>
        <v-card-actions>
          <v-spacer></v-spacer>
          <v-btn color="secondary" @click="dialog = false">Cancel</v-btn>
          <v-btn color="primary" @click="saveAutoFilter">Save</v-btn>
        </v-card-actions>
      </v-card>
    </v-dialog>

    <v-text-field label="Current row" v-model="currentRow" density="compact" @keyup.enter="navigateToRow"
      v-if="columns.length > 0"></v-text-field>

    <!-- <v-checkbox v-model="autofit" label="Autofit columns" density="compact" v-if="columns.length > 0"/> -->

    <v-row v-for="(item, index) in columns" :key="index">
      <v-col cols="12">
        <v-textarea :label="item.columnName + ' - ' + item.header" filled v-model="item.value" rows="1" auto-grow
          density="compact" :prepend-icon="index > 0 ? 'mdi-arrow-up' : '-'"
          @click:prepend="index > 0 && handleMove(index, -1)"
          :prepend-inner-icon="index > 1 ? 'mdi-arrow-collapse-up' : ''"
          @click:prepend-inner="index > 1 && handleMoveToTop(index)"
          :append-icon="index < columns.length - 1 ? 'mdi-arrow-down' : '-'"
          @click:append="index < columns.length - 1 && handleMove(index, 1)"
          :append-inner-icon="item.value ? _currentClipboardText === item.value ? 'mdi-check' : 'mdi-content-copy' : ''"
          @click:append-inner="copyToClipboard(item.value)" @input="onTextChange(index)"></v-textarea>
      </v-col>
    </v-row>
  </v-container>
</template>


<script lang="ts">
import Dexie, { type Table } from 'dexie';

interface IFilter {
  header: string;
  filters: string[];
}

interface IFilterData {
  id?: number; // Auto-incremented primary key
  workbookName: string;
  worksheetName: string;
  title: String;
  filterData: IFilter[];
}

interface IColumnOrder {
  workbookName: string;
  worksheetName: string;
  headers: string[];
}

interface Column {
  position: number;
  columnName: string;
  header: string;
  value: string;
}
const SHOW_ALL = "Show all";

class DbProxy extends Dexie {
  filters!: Table<IFilterData>;
  columnOrders!: Table<IColumnOrder>;

  constructor() {
    super('PivotFromDB');

    this.version(1).stores({
      filters: '++id, [workbookName+worksheetName], [workbookName+worksheetName+title]',
      columnOrders: '[workbookName+worksheetName]'
    });
  }
}
const db = new DbProxy();

export default {
  name: 'App',
  data: function () {
    return {
      currentRow: 0,
      previousOkRow: 0,
      // autofit: true,
      _currentClipboardText: "",

      columns: [] as Column[],

      filters: [] as IFilterData[],
      currentFilter: "",

      dialog: false,
      inputValue: 'Default filter',
      valid: false,

      _context: Excel.RequestContext.prototype,
      _workbook: Excel.Workbook.prototype,
      _worksheet: Excel.Worksheet.prototype,
      emptyFilter: [] as IFilter[],
    };
  },

  async mounted() {
    // const autofitString = localStorage.getItem('autofit');
    // if (autofitString === 'false')
    //   this.autofit = false;

    await Excel.run(async (context: Excel.RequestContext) => {
      this._context = context;
      this._workbook = context.workbook.load(["name"]);

      this.checkForAutoFilter();

      // Add an event handler for the activated sheet event
      this._workbook.worksheets.onActivated.add(this.checkForAutoFilter.bind(this));
    });
  },

  // watch: {
  //   autofit: function (value: boolean) {
  //     localStorage.setItem('autofit', value.toString());
  //   }
  // },

  methods: {

    async handleMove(index: number, dx: number) {
      // Swap the current item with the one above or below element
      const temp: Column = this.columns[index];
      this.columns[index] = this.columns[index + dx];
      this.columns[index + dx] = temp;

      this.saveOrder();
    },

    async handleMoveToTop(index: number) {
      // Move the current item to the top
      const temp: Column = this.columns[index];
      this.columns.splice(index, 1);
      this.columns.unshift(temp);

      this.saveOrder();
    },

    async copyToClipboard(text: string) {
      this._currentClipboardText = text;
      await navigator.clipboard.writeText(text);

      setTimeout(() => {
        this._currentClipboardText = '';
      }, 1500);
    },

    async saveOrder() {
      await db.columnOrders.put({
        workbookName: this._workbook.name,
        worksheetName: this._worksheet.name,
        headers: this.columns.map(column => column.header)
      } as IColumnOrder);
    },

    async updateFilterList(newFilter?: IFilterData) {
      if (newFilter) {
        await db.filters
          .where({ workbookName: this._workbook.name, worksheetName: this._worksheet.name, title: newFilter.title })
          .delete();

        if (newFilter.filterData && newFilter.filterData.length !== 0)
          await db.filters.add(newFilter);
      }
      this.filters = await db.filters
        .where('[workbookName+worksheetName]')
        .equals([this._workbook.name, this._worksheet.name])
        .toArray()
        // Newest filters first
        .then(filters => filters.reverse());

      // Add the default filter -> Show all
      this.filters.unshift({
        workbookName: '',
        worksheetName: '',
        title: SHOW_ALL,
        filterData: this.emptyFilter
      } as IFilterData);
    },

    async saveAutoFilter() {
      this.dialog = false;

      if (!this._workbook || this.inputValue === SHOW_ALL || !this.inputValue)
        return;

      const activeAutoFilter = this._worksheet.autoFilter.load(["criteria"]);
      const activeRange = activeAutoFilter.getRange().load(["values"]);
      await this._context.sync();

      // console.log(activeAutoFilter.criteria);

      const filterData: string[][] = activeAutoFilter.criteria.map((criteria) => {
        if (criteria.criterion1) {
          const result = [String(criteria.criterion1).substring(1)];
          if (criteria.criterion2) {
            result.push(String(criteria.criterion2).substring(1));
          }
          return result;
        } else if (criteria.values && criteria.values.length > 0) {
          return criteria.values.map(value => String(value));
        }
        return [];
      });

      const headers = activeRange.values[0]
      const combinedData = headers.map((header, index) => ({
        header: String(header),
        filters: filterData[index] || []
      })
      // Do not save empty filters
      ).filter(item => item.filters.length > 0);

      const newFilter = {
        title: this.inputValue,
        workbookName: this._workbook.name,
        worksheetName: this._worksheet.name,
        filterData: combinedData
      } as IFilterData;
      this.updateFilterList(newFilter);
    },

    async deleteFilter() {
      if (!this.currentFilter || this.currentFilter === SHOW_ALL)
        return;

      this.updateFilterList({
        title: this.currentFilter,
        workbookName: this._workbook.name,
        worksheetName: this._worksheet.name,
        filterData: []
      } as IFilterData);

      this.currentFilter = "";
    },

    async onApplyChange() {
      const selectedFilter = this.filters.find(filter => filter.title === this.currentFilter);
      if (!selectedFilter)
        return;

      const autoFilterRange = this._worksheet?.autoFilter?.getRange()
      if (!autoFilterRange)
        return

      autoFilterRange.load(["values"]);
      await this._context.sync();

      // Load again after sync
      const autoFilterRange1 = this._worksheet?.autoFilter?.getRange()

      const arrHeaders = autoFilterRange.values[0];
      for (let columnIndex = 0; columnIndex < arrHeaders.length; columnIndex++) {
        const header = arrHeaders[columnIndex] as String;

        const filter = selectedFilter.filterData.find(filter => filter.header === header);
        if (!filter || !filter.filters || filter.filters.length === 0) {
          this._worksheet.autoFilter.clearColumnCriteria(columnIndex)
          continue
        }

        this._worksheet.autoFilter.apply(autoFilterRange1, columnIndex, {
          values: filter.filters,
          operator: Excel.FilterOperator.and,
          filterOn: Excel.FilterOn.values
        });
      }

      await this._context.sync();
    },

    async onTextChange(index: number) {
      if (!this.currentRow)
        return

      const column = this.columns[index];
      const range = this._worksheet.getRange(`${column.columnName}${this.currentRow}`);
      range.values = [[column.value]];

      // if (this.autofit){
      //   range.format.autofitColumns();
      // }
      await this._context.sync();
    },

    async checkForAutoFilter() {
      this.columns = [];

      if (!this._workbook)
        return;
      const activeSheet = this._workbook.worksheets.getActiveWorksheet().load(["name"]);
      const autoFilterRange = activeSheet?.autoFilter?.getRange()?.load(["values", "columnIndex"])
      if (!autoFilterRange)
        return

      await this._context.sync();

      this._worksheet = activeSheet;

      const headerColumns = autoFilterRange.values[0];
      for (let i = 0; i < headerColumns.length; i++) {
        const columnName = this.columnIndexToName(i + autoFilterRange.columnIndex)
        this.columns.push({
          position: i, // <-----Here
          columnName: columnName,
          header: headerColumns[i],
          value: "",
        })
      }

      const prevOrder = await db.columnOrders.get([this._workbook.name, this._worksheet.name]);
      if (prevOrder && prevOrder.headers && prevOrder.headers.length > 0)
        this.columns = this.columns.sort((a, b) => prevOrder.headers.indexOf(a.header) - prevOrder.headers.indexOf(b.header));

      this.emptyFilter = headerColumns.map(header => ({ header: String(header), filters: [] }))
      this.updateFilterList()

      activeSheet.onSelectionChanged.add(this.onSelectionChange)
      if (Office.context.platform === Office.PlatformType.OfficeOnline) {
        this._worksheet.namedSheetViews.enterTemporary();
      }
      await this._context.sync();
    },

    async getCurrentRow(address: string) {
      const autoFilterRange = this._worksheet.autoFilter.getRange().load(["rowIndex"]);
      const currentSelectionRange = this._worksheet.getRange(address);

      const intersection = autoFilterRange.getIntersectionOrNullObject(currentSelectionRange).load(["rowIndex"]);
      await this._context.sync();

      // Set new row if intersection is not null
      if (!intersection.isNullObject) {
        this.previousOkRow = intersection.rowIndex + 1
      }

      return this.previousOkRow
    },

    async navigateToRow() {
      try {
        const activeCell = this._workbook.getActiveCell().load("columnIndex");
        await this._context.sync();

        const columnName = this.columnIndexToName(activeCell.columnIndex);
        this.currentRow = await this.getCurrentRow(`${columnName}${this.currentRow}`)

        const range = this._worksheet.getRange(`${columnName}${this.currentRow}`)
        range.select();
        await this._context.sync();
      } catch (error) {
        console.error('Error navigating to row:', error);
      }
    },

    async onSelectionChange(params: Excel.WorksheetSelectionChangedEventArgs) {
      this.currentRow = await this.getCurrentRow(params.address)
      if (!this.currentRow)
        return

      const autoFilterRange = this._worksheet.autoFilter.getRange().load(["values", "rowIndex"]);
      await this._context.sync();

      const rowValues = autoFilterRange.values[this.currentRow - autoFilterRange.rowIndex - 1]
      // Transpose the row values to the columns in the task pane
      for (let i = 0; i < rowValues.length; i++) {
        this.columns[i].value = rowValues[this.columns[i].position]
      }
    },

    columnIndexToName(index: number) {
      let columnName = "";
      while (index >= 0) {
        columnName = String.fromCharCode((index % 26) + 65) + columnName;
        index = Math.floor(index / 26) - 1;
      }
      return columnName;
    },
  }
}

</script>