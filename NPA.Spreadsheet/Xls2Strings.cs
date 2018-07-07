using System;
using System.Collections;
using System.Collections.Generic;
using NPOI.HSSF.EventUserModel;
using NPOI.HSSF.EventUserModel.DummyRecord;
using NPOI.HSSF.Model;
using NPOI.HSSF.Record;
using NPOI.HSSF.UserModel;
using NPOI.POIFS.FileSystem;

namespace NPA.Spreadsheet
{
    /// <summary>
    /// A XLS -> string processor, that uses the MissingRecordAware
    /// EventModel code to ensure it outputs all columns and rows.
    /// </summary>
    internal class Xls2Strings : HSSFListener
    {
        private readonly POIFSFileSystem _fs;
        public IList<IList<string>> Output =
            new List<IList<string>>();

        private IList<string> _row =
            new List<string>();

        private int _lastRowNumber;

        // Should we output the formula, or the value it has? */
        public bool OutputFormulaValues;

        // For parsing Formulas
        private EventWorkbookBuilder.SheetRecordCollectingListener _workbookBuildingListener;
        private HSSFWorkbook _stubWorkbook;

        // Records we pick up as we process
        private SSTRecord _sstRecord;
        private FormatTrackingHSSFListener _formatListener;
	
        /** So we known which sheet we're on */
        private BoundSheetRecord[] _orderedBsRs;
        private readonly ArrayList _boundSheetRecords = new ArrayList();

        // For handling formulas with string results
        private int _nextRow;
        private bool _outputNextStringRecord;

        /// <summary>
        /// Creates a new XLS -> strings converter
        /// </summary>
        /// <param name="fs">The POIFSFileSystem to process</param>
        public Xls2Strings(POIFSFileSystem fs)
        {
            _fs = fs;
        }

        /// <summary>
        /// Initiates the processing of the XLS file to strings
        /// </summary>
        public void Process()
        {
            var listener = new MissingRecordAwareHSSFListener(this);
            _formatListener = new FormatTrackingHSSFListener(listener);

            var factory = new HSSFEventFactory();
            var request = new HSSFRequest();

            _workbookBuildingListener = new EventWorkbookBuilder.SheetRecordCollectingListener(_formatListener);
            request.AddListenerForAllRecords(_workbookBuildingListener);

            factory.ProcessWorkbookEvents(request, _fs);
        }

        /// <summary>
        /// Main HSSFListener method, processes events, and outputs the
        /// strings as the file is processed
        /// </summary>
        public void ProcessRecord(Record record)
        {
            var thisRow = -1;
            string thisStr = null;

            switch (record.Sid)
            {
                case BoundSheetRecord.sid:
                    _boundSheetRecords.Add(record);
                    break;

                case BOFRecord.sid:
                    var br = (BOFRecord)record;
                    if (br.Type == BOFRecord.TYPE_WORKSHEET)
                    {
                        // Create sub workbook if required
                        if (_workbookBuildingListener != null && _stubWorkbook == null)
                        {
                            _stubWorkbook = _workbookBuildingListener.GetStubHSSFWorkbook();
                        }
				
                        // Output the worksheet name
                        // Works by ordering the BSRs by the location of
                        //  their BOFRecords, and then knowing that we
                        //  process BOFRecords in byte offset order
                        if(_orderedBsRs == null)
                        {
                            _orderedBsRs = BoundSheetRecord.OrderByBofPosition(_boundSheetRecords);
                        }
                    }
                    break;

                case SSTRecord.sid:
                    _sstRecord = (SSTRecord) record;
                    break;

                case BlankRecord.sid:
                    var brec = (BlankRecord) record;
                    thisRow = brec.Row;
                    thisStr = "";
                    break;

                case BoolErrRecord.sid:
                    var berec = (BoolErrRecord) record;
                    thisRow = berec.Row;
                    thisStr = "";
                    break;

                case FormulaRecord.sid:
                    var frec = (FormulaRecord) record;
                    thisRow = frec.Row;
                    if (OutputFormulaValues)
                    {
                        if (double.IsNaN(frec.Value))
                        {
                            // Formula result is a string
                            // This is stored in the next record
                            _outputNextStringRecord = true;
                            _nextRow = frec.Row;
                        }
                        else
                        {
                            thisStr = _formatListener.FormatNumberDateCell(frec);
                        }
                    }
                    else
                    {
                        thisStr = HSSFFormulaParser.ToFormulaString(_stubWorkbook, frec.ParsedExpression);
                    }
                    break;

                case StringRecord.sid:
                    if (_outputNextStringRecord)
                    {
                        // String for formula
                        var srec = (StringRecord)record;
                        thisStr = srec.String;
                        thisRow = _nextRow;
                        _outputNextStringRecord = false;
                    }
                    break;

                case LabelRecord.sid:
                    var lrec = (LabelRecord) record;
                    thisRow = lrec.Row;
                    thisStr = lrec.Value;
                    break;

                case LabelSSTRecord.sid:
                    var lsrec = (LabelSSTRecord) record;

                    thisRow = lsrec.Row;
                    if (_sstRecord == null)
                    {
                        thisStr = "(No SST Record, can't identify string)";
                    }
                    else
                    {
                        thisStr = _sstRecord.GetString(lsrec.SSTIndex).ToString();
                    }
                    break;

                case NoteRecord.sid:
                    var nrec = (NoteRecord) record;
                    thisRow = nrec.Row;
                    thisStr = "";
                    break;

                case NumberRecord.sid:
                    var numrec = (NumberRecord) record;
                    thisRow = numrec.Row;
                    // Format
                    thisStr = _formatListener.FormatNumberDateCell(numrec);
                    break;

                case RKRecord.sid:
                    var rkrec = (RKRecord) record;
                    thisRow = rkrec.Row;
                    thisStr = "";
                    break;
            }

            // Handle new row
            if (thisRow != -1 && thisRow != _lastRowNumber)
            {
                _row = new List<string>();
            }

            // Handle missing column
            if (record is MissingCellDummyRecord)
            {
                var mc = (MissingCellDummyRecord)record;
                thisRow = mc.Row;
                thisStr = "";
            }

            // If we got something to print out, do so
            if (thisStr != null)
            {
                _row.Add(thisStr);
            }

            // Update column and row count
            if (thisRow > -1)
                _lastRowNumber = thisRow;

            // Handle end of row
            if (record is LastCellOfRowDummyRecord)
            {
                // We're onto a new row
                Output.Add(_row);
            }
        }
    }
}