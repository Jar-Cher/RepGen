using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using Xceed.Document.NET;
using Xceed.Words.NET;
//using System.IO.Packaging;

namespace ClassLibrary7
{
    public class ReportGenerator
    {

        private static ReplaceData _replaceData;

        private static DocX _document;

        private static string _inputAddress;

        private static string _outputAddress;

        class ReplaceData
        {
            public Dictionary<string, string> ReplacePatterns { get; }

            public Dictionary<string, string> ReplacePictures { get; }

            public Dictionary<string, string[]> ReplaceTableColumns { get; }

            public Dictionary<string, string[][]> ReplaceTables { get; }

            public Dictionary<string, string[]> ReplaceNumberedLists { get; }

            public Dictionary<string, string[]> ReplaceBulletedLists { get; }

            public ReplaceData(Dictionary<string, string> replacePatterns, Dictionary<string, string> replacePictures, Dictionary<string, string[]> replaceTableColumns, Dictionary<string, string[][]> replaceTables, Dictionary<string, string[]> replaceNumberedLists, Dictionary<string, string[]> replaceBulletedLists)
            {
                ReplacePatterns = replacePatterns;
                ReplacePictures = replacePictures;
                ReplaceTableColumns = replaceTableColumns;
                ReplaceTables = replaceTables;
                ReplaceNumberedLists = replaceNumberedLists;
                ReplaceBulletedLists = replaceBulletedLists;
            }
        }

        private static void ReplaceLists(Dictionary<string, string[]> lists, ListItemType listType)
        {
            if (lists == null)
            {
                return;
            }
            foreach (var i in lists)
            {
                //Console.Write("Gotcha");
                if (!i.Value.Any())
                {
                    continue;
                }
                _document.SaveAs(_outputAddress);
                _document = DocX.Load(_outputAddress);
                List<int> marks = _document.FindAll("{<" + i.Key + ">}");
                int marksId = 0;
                while (marksId < marks.Count)
                {
                    int element = marks[marksId]+10;
                    
                    var newList = _document.AddList(i.Value[0], 0, listType, 1);
                    for (int j = 0; j < i.Value.Length - 1; ++j)
                    {
                        _document.AddListItem(newList, i.Value[j + 1]);
                    }
                    //Console.WriteLine("Got it");
                    try
                    { 
                        _document.InsertList(element, newList);
                    }
                    catch (System.ArgumentOutOfRangeException)
                    {
                    }
                    marks = _document.FindAll("{<" + i.Key + ">}");
                    ++ marksId;
                }
                // Workaround to fix incorrect indexes
                _document.SaveAs(_outputAddress);
                _document = DocX.Load(_outputAddress);
                _document.ReplaceText("{<" + i.Key + ">}", "");
            }
        }

        public static void Replace(string replaceData, string inputAddress = "Input.docx", string outputAddress = "Output.docx")
        {
            _inputAddress = inputAddress;
            _outputAddress = outputAddress;

            try
            {
                _replaceData = JsonConvert.DeserializeObject<ReplaceData>(replaceData);
            }
            catch (JsonReaderException e)
            {
                Console.Write(e);
                return;
            }

            // Load a _document.
            try
            {
                _document = DocX.Load(inputAddress);
            }
            catch (InvalidOperationException e)
            {
                Console.Write(e);
                return;
            }

            // Check if all the replace patterns are used in the loaded _document.
            if ((_document.FindUniqueByPattern(@"{<(.+)>}", RegexOptions.IgnoreCase).Count > 0) && (_replaceData != null))
            {
                // Take care of tables
                if (_replaceData.ReplaceTables != null)
                {
                    foreach (var i in _replaceData.ReplaceTables)
                    {
                        string[][] tableArray = i.Value;
                        //Console.Write(tableArray.Min(row => row.Length));
                        if ((tableArray.GetLength(0) == 0) || (tableArray.Min(row => row.Length) == 0))
                        {
                            continue;
                        }

                        var table = _document.AddTable(tableArray.GetLength(0), tableArray.Max(row => row.Length));

                        for (int j = 0; j < tableArray.GetLength(0); ++j)
                        {
                            for (int k = 0; k < tableArray[j].Length; ++k)
                            {
                                table.Rows[j].Cells[k].Paragraphs[0].Append(i.Value[j][k]);
                            }
                        }

                        _document.ReplaceTextWithObject("{<" + i.Key + ">}", table, false, RegexOptions.IgnoreCase);
                    }
                }

                // Take care of mutable table columns
                if (_replaceData.ReplaceTableColumns != null && _replaceData.ReplaceTableColumns.Any())
                {
                    var mutableTables = _document.Tables.FindAll(tab =>
                        tab.Rows.Find(cell =>
                            cell.Paragraphs.ToList().Find(par => Regex.IsMatch(par.Text, @"{<TableColumns\.(.+)>}")) !=
                            null) != null);

                    for (int i = 0; i < mutableTables.Count; ++i)
                    {
                        int j = 0;
                        while (j < mutableTables[i].RowCount)
                        {
                            var mutableCells = mutableTables[i].Rows[j].Cells.FindAll(cell =>
                                cell.Paragraphs.ToList()
                                    .Find(par => Regex.IsMatch(par.Text, @"{<TableColumns\.(.+)>}")) !=
                                null);

                            if (mutableCells.Any())
                            {
                                //Console.Write("Gotcha");
                                var keysToReplace = _replaceData.ReplaceTableColumns.Keys.ToList().FindAll(key =>
                                    mutableCells.Find(cell =>
                                        cell.Paragraphs.ToList().Find(par => Regex.IsMatch(par.Text, key)) !=
                                        null) != null);

                                int maxLen = 0;
                                for (int k = 0; k < keysToReplace.Count; ++k)
                                {
                                    int newLen = _replaceData.ReplaceTableColumns[keysToReplace[k]].Length;
                                    if (newLen > maxLen)
                                    {
                                        maxLen = newLen;
                                    }

                                    if (newLen == 0)
                                    {
                                        _replaceData.ReplaceTableColumns[keysToReplace[k]] = _replaceData
                                            .ReplaceTableColumns[keysToReplace[k]]
                                            .Concat(new string[] {"{<" + keysToReplace[k] + ">}"}).ToArray();
                                        //Console.Write(_replaceData.ReplaceTableColumns[keysToReplace[k]].Length);
                                    }
                                }

                                for (int k = 0; k < maxLen - 1; ++k)
                                {
                                    //Console.Write(k);
                                    mutableTables[i].InsertRow(mutableTables[i].Rows[j], j + 1, true);


                                    // merge the rest
                                }

                                for (int k = 0; k < maxLen; ++k)
                                {
                                    for (int y = 0; y < keysToReplace.Count; ++y)
                                    {
                                        int cappedIndex =
                                            Math.Min(_replaceData.ReplaceTableColumns[keysToReplace[y]].Length - 1, k);
                                        mutableTables[i].Rows[j].ReplaceText("{<" + keysToReplace[y] + ">}",
                                            _replaceData.ReplaceTableColumns[keysToReplace[y]][cappedIndex], false,
                                            RegexOptions.IgnoreCase);
                                    }
                                    ++j;
                                }
                            }
                            else
                            {
                                ++j;
                            }
                        }
                    }
                }

                // Take care of unifying table cells
                for (int tableId = 0; tableId < _document.Tables.Count; ++tableId)
                {
                    var marks = new Dictionary<int, int>();
                    for (int rowId = 0; rowId < _document.Tables[tableId].Rows.Count; ++rowId)
                    {
                        for (int cellId = 0; cellId < _document.Tables[tableId].Rows[rowId].Cells.Count; ++cellId)
                        {
                            var mark = Regex.Match(string.Join("\n", 
                                _document.Tables[tableId].Rows[rowId].Cells[cellId].Paragraphs.ToList().Select(p => p.Text)), 
                                @"{<ColumnsMerge>}");
                            if (mark.Success)
                            {
                                if (marks.ContainsKey(cellId))
                                {
                                    _document.Tables[tableId].MergeCellsInColumn(cellId, marks[cellId], rowId);
                                    marks.Remove(cellId);
                                }
                                else
                                {
                                    marks[cellId] = rowId;
                                }
                            }
                        }
                    }
                }
                // Workaround to fix incorrect indexes
                _document.SaveAs(_outputAddress);
                _document = DocX.Load(_outputAddress);
                _document.ReplaceText("{<ColumnsMerge>}", "");

                // Take care of lists
                ReplaceLists(_replaceData.ReplaceNumberedLists, ListItemType.Numbered);
                
                ReplaceLists(_replaceData.ReplaceBulletedLists, ListItemType.Bulleted);

                // Take care of pictures
                if (_replaceData.ReplacePictures != null && _replaceData.ReplacePictures.Any())
                {
                    foreach (var i in _replaceData.ReplacePictures)
                    {
                        try
                        {
                            var image = _document.AddImage(i.Value);
                            var picture = image.CreatePicture();
                            // Do the replacement of all the found tags with the specified image and ignore the case when searching for the tags.
                            _document.ReplaceTextWithObject("{<" + i.Key + ">}", picture, false,
                                RegexOptions.IgnoreCase);
                        }
                        catch (FileNotFoundException)
                        {

                        }
                    }
                }

                // Take care of text
                if (_replaceData.ReplacePatterns != null)
                {
                    foreach (var i in _replaceData.ReplacePatterns)
                    {
                        _document.ReplaceText("{<" + i.Key + ">}", i.Value, false, RegexOptions.IgnoreCase);
                    }
                }
                //_document.UpdateFields();
            }
            // Save this _document to disk.
            _document.SaveAs(outputAddress);
        }
        private static string ReplaceString(string findStr)
        {
            
            if (_replaceData.ReplacePatterns.ContainsKey(findStr))
            {
                return _replaceData.ReplacePatterns[findStr];
            }
            return "{<" + findStr + ">}";
        }
    }
}
