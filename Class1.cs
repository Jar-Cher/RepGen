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

        public enum ListType
        {
            Numbered,
            Bulleted
        }

        class ReplaceData
        {
            public Dictionary<string, string> ReplacePatterns { get; }

            public string[] ReplacePictures { get; }

            public Dictionary<string, string[]> ReplaceTableColumns { get; }

            public Dictionary<string, string[][]> ReplaceTables { get; }

            public Dictionary<string, string[]> ReplaceNumberedLists { get; }

            public Dictionary<string, string[]> ReplaceBulletedLists { get; }

            public ReplaceData(Dictionary<string, string> replacePatterns, string[] replacePictures, Dictionary<string, string[]> replaceTableColumns, Dictionary<string, string[][]> replaceTables, Dictionary<string, string[]> replaceNumberedLists, Dictionary<string, string[]> replaceBulletedLists)
            {
                ReplacePatterns = replacePatterns;
                ReplacePictures = replacePictures;
                ReplaceTableColumns = replaceTableColumns;
                ReplaceTables = replaceTables;
                ReplaceNumberedLists = replaceNumberedLists;
                ReplaceBulletedLists = replaceBulletedLists;
            }
        }

        static private void ReplaceLists(Document document, Dictionary<string, string[]> lists, ListItemType listType)
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

                List<int> marks = document.FindAll("{<" + i.Key + ">}");
                marks.ForEach(element =>
                {
                    var newList = document.AddList(i.Value[0], 0, listType, 1);
                    for (int j = 0; j < i.Value.Length - 1; ++j)
                    {
                        document.AddListItem(newList, i.Value[j + 1]);
                    }

                    try
                    {
                        document.InsertList(element, newList);
                    }
                    catch (System.ArgumentOutOfRangeException)
                    {

                    }
                });
                document.ReplaceText("{<" + i.Key + ">}", "");
            }
        }

        public static void Replace(string replaceData, string inputAddress = "Input.docx", string outputAddress = "Output.docx")
        {
            _replaceData = JsonConvert.DeserializeObject<ReplaceData>(replaceData);

            // Load a document.
            var document = DocX.Load(inputAddress);
            // Check if all the replace patterns are used in the loaded document.
            if ((document.FindUniqueByPattern(@"{<(.+)>}", RegexOptions.IgnoreCase).Count > 0) && (_replaceData != null))
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

                        var table = document.AddTable(tableArray.GetLength(0), tableArray.Max(row => row.Length));

                        for (int j = 0; j < tableArray.GetLength(0); ++j)
                        {
                            for (int k = 0; k < tableArray[j].Length; ++k)
                            {
                                table.Rows[j].Cells[k].Paragraphs[0].Append(i.Value[j][k]);
                            }
                        }

                        document.ReplaceTextWithObject("{<" + i.Key + ">}", table, false,
                            RegexOptions.IgnoreCase);
                    }
                }

                // Take care of mutable table columns
                var mutableTables = document.Tables.FindAll(tab => 
                    tab.Rows.Find(cell => cell.Paragraphs.ToList().Find(par => Regex.IsMatch(par.Text, @"{<TableColumns\.(.+)>}")) != null) != null);

                for (int i = 0; i < mutableTables.Count; ++i)
                {
                    int j = 0;
                    while (j < mutableTables[i].RowCount)
                    {
                        var mutableCells = mutableTables[i].Rows[j].Cells.FindAll(cell =>
                            cell.Paragraphs.ToList().Find(par => Regex.IsMatch(par.Text, @"{<TableColumns\.(.+)>}")) !=
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
                                    _replaceData.ReplaceTableColumns[keysToReplace[k]] = _replaceData.ReplaceTableColumns[keysToReplace[k]].Concat(new string[] { "{<" + keysToReplace[k] + ">}" }).ToArray();
                                    //Console.Write(_replaceData.ReplaceTableColumns[keysToReplace[k]].Length);
                                }
                            }

                            for (int k = 0; k < maxLen-1; ++k)
                            {
                                //Console.Write(k);
                                mutableTables[i].InsertRow(mutableTables[i].Rows[j], j+1, true);
                                
                                
                                // merge the rest
                            }

                            for (int k = 0; k < maxLen; ++k)
                            {
                                for (int y = 0; y < keysToReplace.Count; ++y)
                                {
                                    int cappedIndex = Math.Min(_replaceData.ReplaceTableColumns[keysToReplace[y]].Length - 1, k);
                                    mutableTables[i].Rows[j].ReplaceText("{<" + keysToReplace[y] + ">}", _replaceData.ReplaceTableColumns[keysToReplace[y]][cappedIndex], false, RegexOptions.IgnoreCase);
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
                
                // Take care of lists
                ReplaceLists(document, _replaceData.ReplaceNumberedLists, ListItemType.Numbered);
                // Workaround to fix incorrect indexes
                if (_replaceData.ReplaceNumberedLists != null)
                {
                    document.SaveAs(outputAddress);
                    document = DocX.Load(outputAddress);

                    foreach (var i in _replaceData.ReplaceNumberedLists)
                    {
                        document.ReplaceText("{<" + i.Key + ">}", "");
                    }
                }
                ReplaceLists(document, _replaceData.ReplaceBulletedLists, ListItemType.Bulleted);
                if (_replaceData.ReplaceBulletedLists != null)
                {
                    document.SaveAs(outputAddress);
                    document = DocX.Load(outputAddress);

                    foreach (var i in _replaceData.ReplaceBulletedLists)
                    {
                        document.ReplaceText("{<" + i.Key + ">}", "");
                    }
                }

                // Take care of pictures
                for (int i = 0; i < _replaceData.ReplacePictures.Length; ++i)
                {
                    try
                    {
                        var image = document.AddImage(_replaceData.ReplacePictures[i]);
                        var picture = image.CreatePicture();
                        // Do the replacement of all the found tags with the specified image and ignore the case when searching for the tags.
                        document.ReplaceTextWithObject("{<" + _replaceData.ReplacePictures[i] + ">}", picture, false,
                            RegexOptions.IgnoreCase);
                    }
                    catch (System.IO.FileNotFoundException)
                    {
                        
                    }
                }

                // Take care of text
                for (int i = 0; i < _replaceData.ReplacePatterns.Count; ++i)
                {
                    document.ReplaceText("{<(.+)>}", ReportGenerator.ReplaceString, false, RegexOptions.IgnoreCase);
                }
                // Save this document to disk.
                //document.UpdateFields();
            }
            document.SaveAs(outputAddress);
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
