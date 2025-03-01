
﻿
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using OfficeOpenXml;
using System.Diagnostics;

namespace ALGOOO
{
    

    public class DisjointSet<T>
    {
        private Dictionary<T, T> parent;
        private Dictionary<T, int> rank;

        public DisjointSet(IEnumerable<T> elements)
        {
            parent = new Dictionary<T, T>();
            rank = new Dictionary<T, int>();
            foreach (var element in elements)
            {
                parent[element] = element;
                rank[element] = 0;
            }
        }
        public T Find(T element)
        {
            if (!parent.ContainsKey(element))
                throw new KeyNotFoundException("Element not found in the disjoint set.");
            if (!element.Equals(parent[element]))
                parent[element] = Find(parent[element]);
            return parent[element];
        }
        public void Union(T element1, T element2)
        {
            var root1 = Find(element1);
            var root2 = Find(element2);
            if (!root1.Equals(root2))
            {
                if (rank[root1] < rank[root2])
                    parent[root1] = root2;
                else if (rank[root1] > rank[root2])
                    parent[root2] = root1;
                else
                {
                    parent[root2] = root1;
                    rank[root1]++;
                }
            }
        }
    }
    internal class Program
    {
        static Dictionary<string, string> HASHH = new Dictionary<string, string>();
        static Dictionary<string, string> HASHH_mst = new Dictionary<string, string>();
        static Dictionary<string, Tuple<string, string>> HASHH_unique = new Dictionary<string, Tuple<string, string>>();
        //MAIN FUNCTION
        static void Main()
        {
            Stopwatch totalWatch = new Stopwatch();
            totalWatch.Start();
            OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            //----------------------------------------------------------------------------------------------------------------------
            //------------------------------------------------------Load Data--------------------------------------------------------
            Stopwatch loadDataWatch = new Stopwatch();
            loadDataWatch.Start();
            Dictionary<Tuple<string, string>, Tuple<double, double, double>> matchingPairs = LOAD_DATA();
            loadDataWatch.Stop();
            Console.WriteLine($"LOAD_DATA executed in {loadDataWatch.ElapsedMilliseconds} ms");
            //-----------------------------------------------------------------------------------------------------------------------
            //-----------------------------------------------Create Connected Components---------------------------------------------
            Stopwatch createComponentsWatch = new Stopwatch();
            createComponentsWatch.Start();
            List<List<string>> components = CreateConnectedComponents(matchingPairs);
            createComponentsWatch.Stop();
            Console.WriteLine($"CreateConnectedComponents executed in {createComponentsWatch.ElapsedMilliseconds} ms");
            //------------------------------------------------------------------------------------------------------------------------
            //-----------------------------------------------Calculate Average Weight and Count--------------------------------------- 
            Stopwatch calculateWeightWatch = new Stopwatch();
            calculateWeightWatch.Start();
            List<Tuple<List<string>, double, int>> Connected_Components_with_Average_Weight =
            new List<Tuple<List<string>, double, int>>();
            foreach (List<string> component in components)
            {
                int count;
                double averageWeight;
                (averageWeight, count) = CalculateAverageWeight(component, matchingPairs);
                Tuple<List<string>, double, int> Connected_ = Tuple.Create(component, averageWeight, count);
                Connected_Components_with_Average_Weight.Add(Connected_);
            }
            calculateWeightWatch.Stop();
            Console.WriteLine($"CalculateAverageWeight executed in {calculateWeightWatch.ElapsedMilliseconds} ms");
            //----------------------------------------------------------------------------------------------------------------------------
            //--------------------------------------------Sorting First OutPut--------------------------------------------------------------
            List<Tuple<List<string>, double, int>> sortedList = Connected_Components_with_Average_Weight.OrderByDescending(t => t.Item2).ToList();
            List<Tuple<List<string>, double, int>> Listt = new List<Tuple<List<string>, double, int>>();
            foreach (var list in sortedList)
            {
                List<string> hi = list.Item1;
                List<string> sortedHi = hi.OrderBy(x => int.Parse(x)).ToList();
                Tuple<List<string>, double, int> Connected_ = Tuple.Create(sortedHi, list.Item2, list.Item3);
                Listt.Add(Connected_);
            }
            //-------------------------------------------------------------------------------------------------------------------------------
            //-------------------------------------------------To Excel Sheet----------------------------------------------------------------
            Stopwatch ListToExcel = new Stopwatch();
            ListToExcel.Start();
            ConvertListToExcel(Listt);
            ListToExcel.Stop();
            Console.WriteLine($"Convert List To Excel executed in {ListToExcel.ElapsedMilliseconds} ms");
            //-------------------------------------------------------------------------------------------------------------------------------
            //----------------------------------------------Convert Data To Use In MST-------------------------------------------------------
            Stopwatch calculate = new Stopwatch();
            calculate.Start();
            List<Dictionary<Tuple<string, string>, Tuple<double, double, double>>> componentEdges =
            GetEdgesInComponent(components, matchingPairs);
            calculate.Stop();
            Console.WriteLine($"CalculateAverageWeight executed in {calculate.ElapsedMilliseconds} ms");
            //--------------------------------------------------------------------------------------------------------------------------------
            //--------------------------------------------------------MST---------------------------------------------------------------------
            Stopwatch findMSTWatch = new Stopwatch();
            findMSTWatch.Start();
            List<Dictionary<Tuple<string, string>, double>> data1;
            data1 = FindMaxSpanningTree(components, componentEdges);
            findMSTWatch.Stop();
            Console.WriteLine($"FindMaxSpanningTree executed in {findMSTWatch.ElapsedMilliseconds} ms");
            


            // Prim's MST

            Stopwatch findMSTPrimWatch = new Stopwatch();
            findMSTPrimWatch.Start();
            List<Dictionary<Tuple<string, string>, double>> dataPrim = FindMSTUsingPrim(components, componentEdges);
            findMSTPrimWatch.Stop();
            Console.WriteLine($"FindMSTUsingPrim executed in {findMSTPrimWatch.ElapsedMilliseconds} ms");




            //--------------------------------------------------------------------------------------------------------------------------------
            //--------------------------------------------------Sorrt Second Output-----------------------------------------------------------
            List<Tuple<Dictionary<Tuple<string, string>, double>, double>> dataaa =
            new List<Tuple<Dictionary<Tuple<string, string>, double>, double>>(data1.Count);
            //---------sort el data mn gwa--------
            for (int i = 0; i < data1.Count; i++)
            {
                var dictionary = data1[i];
                var sortedKeyValuePairs = dictionary.OrderByDescending(kvp => kvp.Value).ToList();
                Dictionary<Tuple<string, string>, double> sortedDictionary =
                sortedKeyValuePairs.ToDictionary(kvp => kvp.Key, kvp => kvp.Value);
            }
            //----------sort el componants nfsha--------------
            for (int i = 0; i < componentEdges.Count; i++)
            {
                var dictionary = componentEdges[i];
                var sortedKeyValuePairs = dictionary.OrderByDescending(kvp => kvp.Value.Item3).ToList();
                Dictionary<Tuple<string, string>, Tuple<double, double, double>> sortedDictionary =
                    sortedKeyValuePairs.ToDictionary(kvp => kvp.Key, kvp => kvp.Value);
                //--------------------------------------------------------------------------

                double sum = 0;
                double average = 0;
                int counter = 0;
                string[] parts1, parts2;
                List<Tuple<string, string>> keysToRemove = new List<Tuple<string, string>>();

                foreach (var kvp in sortedDictionary)
                {
                    string hashing1 = kvp.Key.Item1 + kvp.Value.Item3.ToString();
                    string hashing2 = kvp.Key.Item2 + kvp.Value.Item3.ToString();
                    string node1 = "";
                    string node2 = "";
                    bool flag = false;

                    try
                    {
                        string node5 = HASHH_mst[hashing1];
                        string node6 = HASHH_mst[hashing2];
                        flag = true;
                    }
                    catch
                    {
                        flag = false;
                    }

                    node1 = HASHH[hashing1];
                    node2 = HASHH[hashing2];

                    parts1 = node1.Split('/');
                    parts2 = node2.Split('/');

                    // Take the second-to-last part, which should contain the ID
                    string similarityPart1 = Regex.Replace(parts1[parts1.Length - 1], "[^0-9]", "");
                    string similarityPart2 = Regex.Replace(parts2[parts2.Length - 1], "[^0-9]", "");
                    double weightOne = double.Parse(similarityPart1);
                    double weightTwo = double.Parse(similarityPart2);
                    counter += 2;
                    sum += weightOne + weightTwo;

                    if (!flag)
                    {
                        keysToRemove.Add(kvp.Key);
                    }
                }

                // Remove keys from sortedDictionary
                foreach (var keyToRemove in keysToRemove)
                {
                    sortedDictionary.Remove(keyToRemove);
                }

                average = sum / counter;

                Dictionary<Tuple<string, string>, double> khello = new Dictionary<Tuple<string, string>, double>();
                foreach (var x in sortedDictionary)
                {
                    khello[x.Key] = x.Value.Item3;
                }

                Tuple<Dictionary<Tuple<string, string>, double>, double> componentData = Tuple.Create(khello, average);
                dataaa.Add(componentData);
            }
            dataaa = dataaa.OrderByDescending(item => item.Item2).ToList();
            //---------------------------------------------------------------------------------------------------------------------------------
            //---------------------------------------------------To Excel Sheet2---------------------------------------------------------------
            Stopwatch convertExcelWatch = new Stopwatch();
            convertExcelWatch.Start();
            ConvertDictionaryToExcel(dataaa);
            convertExcelWatch.Stop();
            Console.WriteLine($"ConvertDictionaryToExcel executed in {convertExcelWatch.ElapsedMilliseconds} ms");
            //----------------------------------------------------------------------------------------------------------------------------------
            totalWatch.Stop();
            Console.WriteLine($"Total execution time: {totalWatch.ElapsedMilliseconds} ms");
            Console.ReadLine();
        }
        //===============================================================================================================================================
        //---------------------------------------------------Load Data-----------------------------------------------------------------------------------
        //Dina
        private static Dictionary<Tuple<string, string>, Tuple<double, double, double>> LOAD_DATA()
        {
            string path = @"D:\GAM3A\ALGOOO\PROJECT\[3] Plagiarism Validation\Test Cases\Complete\Hard\1-Input.xlsx";
            var matchingPairs = new Dictionary<Tuple<string, string>, Tuple<double, double, double>>();

            FileInfo fileInfo = new FileInfo(path);
            using (var package = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // First worksheet
                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    if (worksheet.Cells[row, 1].Value == null || worksheet.Cells[row, 2].Value == null)
                        break;
                    string file1Path = worksheet.Cells[row, 1].Text;
                    string file2Path = worksheet.Cells[row, 2].Text;
                    string Linesmatch = worksheet.Cells[row, 3].Text;
                    string[] parts1 = file1Path.Split('/');
                    string[] parts2 = file2Path.Split('/');
                    string similarityPart1 = Regex.Replace(parts1[parts1.Length - 1], "[^0-9]", "");
                    string similarityPart2 = Regex.Replace(parts2[parts2.Length - 1], "[^0-9]", "");
                    string idPart1 = Regex.Replace(parts1[parts1.Length - 2], "[^0-9]", "");
                    string idPart2 = Regex.Replace(parts2[parts2.Length - 2], "[^0-9]", "");

                    string hashing1 = idPart1 + Linesmatch;
                    string hashing2 = idPart2 + Linesmatch;
                    string hashing_unique1 = idPart1 + idPart2 + Linesmatch;

                    HASHH[hashing1] = file1Path;
                    HASHH[hashing2] = file2Path;
                    Tuple<string, string> rawan = Tuple.Create(file1Path, file2Path);
                    HASHH_unique[hashing_unique1] = rawan;

                    double weightOne = double.Parse(similarityPart1);
                    double weightTwo = double.Parse(similarityPart2);

                    Tuple<string, string> pairKey = Tuple.Create(idPart1, idPart2);
                    Tuple<double, double, double> weights = Tuple.Create(weightOne, weightTwo, double.Parse(Linesmatch));
                    matchingPairs.Add(pairKey, weights);
                }
            }
            return matchingPairs;
        }


        //================================================================================================================================================
        //------------------------------------GET CONNECTED COMPONENTS FROM THE GRAPH---------------------------------------------------------------------
        //Ziad
        public static List<List<string>> CreateConnectedComponents(Dictionary<Tuple<string, string>, Tuple<double, double, double>> edges)
        {
            Dictionary<string, List<string>> adjacencyList = new Dictionary<string, List<string>>();
            HashSet<string> visitedNodes = new HashSet<string>();
            List<List<string>> components = new List<List<string>>();

            // Build adjacency list
            foreach (var edge in edges)
            {
                string node1 = edge.Key.Item1;
                string node2 = edge.Key.Item2;

                if (!adjacencyList.ContainsKey(node1))
                {
                    adjacencyList[node1] = new List<string>();
                }
                adjacencyList[node1].Add(node2);

                if (!adjacencyList.ContainsKey(node2))
                {
                    adjacencyList[node2] = new List<string>();
                }
                adjacencyList[node2].Add(node1);
            }

            // Depth-First Search to find and mark all connected components
            void DFS(string node, List<string> component)
            {
                Stack<string> stack = new Stack<string>();
                stack.Push(node);
                visitedNodes.Add(node);

                while (stack.Count > 0)
                {
                    string current = stack.Pop();
                    component.Add(current);

                    foreach (string neighbor in adjacencyList[current])
                    {
                        if (!visitedNodes.Contains(neighbor))
                        {
                            visitedNodes.Add(neighbor);
                            stack.Push(neighbor);
                        }
                    }
                }
            }

            // Initialize DFS for components not yet visited
            foreach (string node in adjacencyList.Keys)
            {
                if (!visitedNodes.Contains(node))
                {
                    List<string> component = new List<string>();
                    DFS(node, component);
                    components.Add(component);
                }
            }

            return components;
        }
        //================================================================================================================================================
        //--------------------------------------------Calculate Average Weight----------------------------------------------------------------------------
        //Rawan 
        static (double, int) CalculateAverageWeight(List<string> component, Dictionary<Tuple<string, string>, Tuple<double, double, double>> graph)
        {
            // Convert the component list to a set for faster lookups
            HashSet<string> componentNodes = new HashSet<string>(component);
            double totalWeight = 0.0;
            int edgeCount = 0;

            // Iterate only over the subset of the graph that is relevant to the component
            foreach (var edge in graph)
            {
                string node1 = edge.Key.Item1;
                string node2 = edge.Key.Item2;

                // Check if both nodes are in the component
                if (componentNodes.Contains(node1) && componentNodes.Contains(node2))
                {
                    totalWeight += (edge.Value.Item1 + edge.Value.Item2);
                    edgeCount++;
                }
            }

            // Calculate the average, accounting for the double counting of weight in bidirectional graphs
            return edgeCount > 0 ? (totalWeight / (edgeCount * 2), component.Count) : (0.0, 0);
        }
        //================================================================================================================================================
        //-----------------------------------------CONVERT THE GRAPH TO LIST OF GRAPH---------------------------------------------------------------------
        //Kello
        static List<Dictionary<Tuple<string, string>, Tuple<double, double, double>>> GetEdgesInComponent(
        List<List<string>> connectedComponents, Dictionary<Tuple<string, string>, Tuple<double, double, double>> edges)
        {
            // Initialize a list of dictionaries to store edges for each component
            List<Dictionary<Tuple<string, string>, Tuple<double, double, double>>> componentEdgesList =
                new List<Dictionary<Tuple<string, string>, Tuple<double, double, double>>>(connectedComponents.Count);

            // Create a dictionary to map each node to its component index
            Dictionary<string, int> nodeComponentIndex = new Dictionary<string, int>();

            // Fill the nodeComponentIndex dictionary with the component index of each node
            for (int i = 0; i < connectedComponents.Count; i++)
            {
                foreach (var node in connectedComponents[i])
                {
                    nodeComponentIndex[node] = i;
                }

                // Initialize an empty dictionary for the current component
                componentEdgesList.Add(new Dictionary<Tuple<string, string>, Tuple<double, double, double>>());
            }

            // Iterate over each edge and add it to the corresponding component's dictionary
            foreach (var edge in edges)
            {
                string node1 = edge.Key.Item1;
                string node2 = edge.Key.Item2;

                // Check if both nodes of the edge belong to the same component
                if (nodeComponentIndex.TryGetValue(node1, out int componentIndex1) &&
                    nodeComponentIndex.TryGetValue(node2, out int componentIndex2) &&
                    componentIndex1 == componentIndex2)
                {
                    // Add the edge to the dictionary of the corresponding component
                    componentEdgesList[componentIndex1][edge.Key] = edge.Value;
                }
            }

            return componentEdgesList;
        }
        //=================================================================================================================================================
        //-----------------------------------------------GET THE MST FOR THE GRAPH-------------------------------------------------------------------------
        //--->Ali Mohamed
        public static List<Dictionary<Tuple<string, string>, double>> FindMaxSpanningTree
        (List<List<string>> connectedComponents, List<Dictionary<Tuple<string, string>, Tuple<double, double, double>>> connectedNodes)
        {
            List<Dictionary<Tuple<string, string>, double>> mstEdges = new List<Dictionary<Tuple<string, string>, double>>();

            foreach (var component in connectedComponents)
            {
                Dictionary<Tuple<string, string>, Tuple<double, double, double>> edges =
                connectedNodes.FirstOrDefault(nodes => nodes.Keys.Select(k => k.Item1).Intersect(component).Any());

                if (edges == null)
                {
                    continue;
                }

                // Create a list of edges from the component's dictionary
                List<Tuple<Tuple<string, string>, Tuple<double, double, double>>> edgeList = edges.Select(kvp => Tuple.Create(kvp.Key, kvp.Value)).ToList();
                // Sort the edges in descending order of weights
                edgeList.Sort((x, y) =>
                {
                    int comparison = Math.Max(y.Item2.Item1, y.Item2.Item2).CompareTo(Math.Max(x.Item2.Item1, x.Item2.Item2));
                    if (comparison == 0)
                    {
                        comparison = y.Item2.Item3.CompareTo(x.Item2.Item3);
                    }
                    return comparison;
                });

                // Initialize a disjoint set to track the connected components
                DisjointSet<string> disjointSet = new DisjointSet<string>(component);

                // Create a dictionary to store the MST edges for the current component
                Dictionary<Tuple<string, string>, double> mstComponentEdges = new Dictionary<Tuple<string, string>, double>();

                foreach (var edge in edgeList)
                {
                    string node1 = edge.Item1.Item1;
                    string node2 = edge.Item1.Item2;

                    if (disjointSet.Find(node1) != disjointSet.Find(node2))
                    {
                        // Add the edge to the MST component
                        mstComponentEdges[edge.Item1] = edge.Item2.Item3;
                        HASHH_mst[node1 + edge.Item2.Item3] = HASHH[node1 + edge.Item2.Item3];
                        HASHH_mst[node2 + edge.Item2.Item3] = HASHH[node2 + edge.Item2.Item3];

                        // Union the two sets
                        disjointSet.Union(node1, node2);
                    }
                }

                // Add the MST component edges to the overall MST edges
                mstEdges.Add(mstComponentEdges);
            }

            return mstEdges;
        }

        //--------------------------------------------------------------------------------------------------///
        //-----------------------------BONUUUUUUUUUUUUUUUUUUUUUUUUSSSSSSSSSSSSSSSSSSSS----------------------///

        public static List<Dictionary<Tuple<string, string>, double>> FindMSTUsingPrim(List<List<string>> connectedComponents, List<Dictionary<Tuple<string, string>, Tuple<double, double, double>>> connectedNodes)
        {
            List<Dictionary<Tuple<string, string>, double>> mstEdges = new List<Dictionary<Tuple<string, string>, double>>();

            foreach (var componentEdges in connectedNodes)
            {
                var keys = componentEdges.Keys.ToList();
                if (keys.Count == 0) continue;
                string startNode = keys[0].Item1;
                var priorityQueue = new SortedSet<Tuple<double, Tuple<string, string>>>(Comparer<Tuple<double, Tuple<string, string>>>.Create((x, y) => x.Item1.CompareTo(y.Item1) != 0 ? x.Item1.CompareTo(y.Item1) : x.Item2.Item1.CompareTo(y.Item2.Item1)));
                var inMST = new HashSet<string>();
                var mstComponentEdges = new Dictionary<Tuple<string, string>, double>();
                inMST.Add(startNode);

                AddEdges(startNode, inMST, priorityQueue, componentEdges);

                while (priorityQueue.Count > 0)
                {
                    var edge = priorityQueue.Min;
                    priorityQueue.Remove(edge);
                    string toNode = edge.Item2.Item2;
                    if (inMST.Contains(toNode)) continue;
                    inMST.Add(toNode);
                    mstComponentEdges.Add(edge.Item2, edge.Item1);
                    AddEdges(toNode, inMST, priorityQueue, componentEdges);
                }
                mstEdges.Add(mstComponentEdges);
                //Console.WriteLine($"MST for component with start node {startNode} has {mstComponentEdges.Count} edges.");
            }
            return mstEdges;
        }

        private static void AddEdges(string node, HashSet<string> inMST, SortedSet<Tuple<double, Tuple<string, string>>> priorityQueue, Dictionary<Tuple<string, string>, Tuple<double, double, double>> edges)
        {
            foreach (var edge in edges)
            {
                if (edge.Key.Item1 == node && !inMST.Contains(edge.Key.Item2))
                {
                    priorityQueue.Add(Tuple.Create(edge.Value.Item3, edge.Key));
                }
                else if (edge.Key.Item2 == node && !inMST.Contains(edge.Key.Item1))
                {
                    priorityQueue.Add(Tuple.Create(edge.Value.Item3, Tuple.Create(edge.Key.Item2, edge.Key.Item1)));
                }
            }
        }




        //=================================================================================================================================================
        //-------------------------------------------------RETURN RESULTS TO EXCEL FILE--------------------------------------------------------------------
        //Omar
        public static void ConvertDictionaryToExcel(List<Tuple<Dictionary<Tuple<string, string>, double>, double>> data)
        {
            if (data.Count == 0)
            {
                Console.WriteLine("No data to write to Excel.");
                return;
            }
            string filePath = @"D:\Output1000.xlsx";
            FileInfo fileInfo = new FileInfo(filePath);

            // Delete existing file to ensure a fresh workbook
            if (fileInfo.Exists)
            {
                fileInfo.Delete();
            }

            using (var package = new ExcelPackage(fileInfo))
            {
                // Create a worksheet
                var worksheet = package.Workbook.Worksheets.Add("Data");

                // Setting up headers
                worksheet.Cells["A1"].Value = "First String";
                worksheet.Cells["B1"].Value = "Second String";
                worksheet.Cells["C1"].Value = "Double Value";

                int row = 2;
                List<object[]> rows = new List<object[]>();

                foreach (var item in data)
                {
                    foreach (var kvp in item.Item1)
                    {
                        string key1 = kvp.Key.Item1 + kvp.Key.Item2 + kvp.Value.ToString();

                        if (!HASHH_unique.ContainsKey(key1))
                        {
                            Console.WriteLine($"Missing data for keys: {key1}");
                            continue;
                        }

                        // Prepare the row data
                        string value1 = HASHH_unique[key1].Item1;
                        string value2 = HASHH_unique[key1].Item2;
                        double value3 = kvp.Value;

                        // Collect row data for batch insertion
                        rows.Add(new object[] { value1, value2, value3 });
                    }
                }

                // Batch load all collected rows into the worksheet
                if (rows.Count > 0)
                {
                    worksheet.Cells[row, 1, row + rows.Count - 1, 3].LoadFromArrays(rows.ToArray());
                }

                // Apply hyperlinks and styles after all rows are populated
                for (int i = 0; i < rows.Count; i++)
                {
                    var cell1 = worksheet.Cells[row + i, 1];
                    var cell2 = worksheet.Cells[row + i, 2];

                    // Check if the cell has text and does not already contain a hyperlink
                    if (Uri.TryCreate(rows[i][0]?.ToString(), UriKind.Absolute, out Uri uri1) && cell1.Hyperlink == null)
                    {
                        cell1.Hyperlink = uri1;
                        cell1.Style.Font.UnderLine = true;
                        cell1.Style.Font.Color.SetColor(System.Drawing.Color.Blue);
                    }

                    if (Uri.TryCreate(rows[i][1]?.ToString(), UriKind.Absolute, out Uri uri2) && cell2.Hyperlink == null)
                    {
                        cell2.Hyperlink = uri2;
                        cell2.Style.Font.UnderLine = true;
                        cell2.Style.Font.Color.SetColor(System.Drawing.Color.Blue);
                    }
                }

                // Save the changes to a new file
                package.Save();
                Console.WriteLine("Excel file created or overwritten at: " + filePath);
            }
        }


        //=================================================================================================================================================
        //-------------------------------------------------RETURN RESULTS TO EXCEL FILE 2--------------------------------------------------------------------
        public static void ConvertListToExcel(List<Tuple<List<string>, double, int>> data)
        {
            string filePath = @"D:\Output1.xlsx";
            FileInfo fileInfo = new FileInfo(filePath);

            // If the file exists and we want to overwrite it, we can delete it beforehand
            if (fileInfo.Exists)
            {
                fileInfo.Delete(); // Delete the existing file to ensure a fresh workbook
            }

            // Using the package with a new FileInfo object, which will create a new file
            using (var package = new ExcelPackage(fileInfo))
            {
                // Create a worksheet
                string sheetName = "Data";
                var worksheet = package.Workbook.Worksheets.Add(sheetName);

                // Setting up headers

                worksheet.Cells["A1"].Value = "Counter";
                worksheet.Cells["B1"].Value = "Component";
                worksheet.Cells["C1"].Value = "Average Similarity";
                worksheet.Cells["D1"].Value = "Count";

                int row = 2;
                int i = 1;
                foreach (var item in data)
                {
                    List<string> component = item.Item1;
                    double averageSimilarity = item.Item2;
                    int count = item.Item3;

                    // Assign values to cells
                    worksheet.Cells[row, 2].Value = string.Join(", ", component);
                    worksheet.Cells[row, 3].Value = Math.Round(averageSimilarity, 1);
                    worksheet.Cells[row, 4].Value = count;
                    worksheet.Cells[row, 1].Value = i;
                    row++;
                    i++;
                }

                // Save the changes to a new file
                package.Save();
                Console.WriteLine("Excel file created or overwritten at: " + filePath);
            }
        }
    }
}