using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml; //EPPlus

namespace LocalSearch3D
{
    class Program
    {
        static void Main(string[] args)
        {
            Random rnd = new Random();
            List<(double[], double)> generatedPoints = new List<(double[], double)>();

            // Area
            double minX = -10.0;
            double maxX = 10.0;
            double minY = -10.0;
            double maxY = 10.0;
            double minZ = -10.0;
            double maxZ = 10.0;

            // Multiplier
            double multiplier;

            // Prompt user for multiplier input
            do
            {
                Console.Write("Enter a multiplier for minValue and maxValue: ");
            } while (!double.TryParse(Console.ReadLine(), out multiplier));

            // Modify minValue and maxValue
            double minValue = -0.8 * multiplier;
            double maxValue = 0.8 * multiplier;

            // Starting point
            double[] startingPoint = { -10.0, -10.0, -10.0 };
            double[] desiredPoint = { 0.0, 0.0, 0.0 };
            double[] currentPos = { startingPoint[0], startingPoint[1], startingPoint[2] };

            // Simulated Annealing Parameters
            double initialTemperature = 100.0;
            double coolingRate = 0.95;
            double temperature = initialTemperature;

            // Find Neighbor
            double[] findNeighbour(double[] coordinates)
            {
                double[] neighbourPoint = { Math.Round(coordinates[0], 4), Math.Round(coordinates[1], 4), Math.Round(coordinates[2], 4) };
                return neighbourPoint;
            }

            // Local Search with Simulated Annealing
            double[] localSearch(double[] coordinates, double temperature)
            {
                double[] randomDot = randomPoint(coordinates);
                double distanceRandom = calculateDistance(randomDot, desiredPoint);
                double distanceCoordinates = calculateDistance(coordinates, desiredPoint);

                // Determine whether to accept a worse solution based on the temperature
                if (distanceRandom < distanceCoordinates || rnd.NextDouble() < Math.Exp(-(distanceRandom - distanceCoordinates) / temperature))
                {
                    coordinates = randomDot;
                }

                return coordinates;
            }

            // Calculate Distance
            double calculateDistance(double[] coordinates1, double[] aim)
            {
                double distance = Math.Sqrt(Math.Pow(coordinates1[0] - aim[0], 2) + Math.Pow(coordinates1[1] - aim[1], 2) + Math.Pow(coordinates1[2] - aim[2], 2));
                return distance;
            }

            // Random Point
            double[] randomPoint(double[] coordinates)
            {
                double[] randomDoubles = {
                    Math.Round(rnd.NextDouble() * (maxValue - minValue) + minValue, 5),
                    Math.Round(rnd.NextDouble() * (maxValue - minValue) + minValue, 5),
                    Math.Round(rnd.NextDouble() * (maxValue - minValue) + minValue, 5)
                };
                double[] coordinatesAdded = { coordinates[0] + randomDoubles[0], coordinates[1] + randomDoubles[1], coordinates[2] + randomDoubles[2] };
                bool check = checkCondition(coordinatesAdded);
                if (check)
                    coordinates = findNeighbour(coordinatesAdded);
                return coordinates;
            }

            // Check Conditions
            bool checkCondition(double[] coordinates)
            {
                return coordinates[0] > minX && coordinates[0] < maxX &&
                       coordinates[1] > minY && coordinates[1] < maxY &&
                       coordinates[2] > minZ && coordinates[2] < maxZ;
            }

            // Body
            double[] nearestPointAchieved = { currentPos[0], currentPos[1], currentPos[2] };
            for (int i = 0; i < 1000; i++)
            {
                currentPos = localSearch(currentPos, temperature);
                double distance = calculateDistance(currentPos, desiredPoint);
                generatedPoints.Add((new double[] { currentPos[0], currentPos[1], currentPos[2] }, distance));

                // Update temperature
                temperature *= coolingRate;

                if (calculateDistance(nearestPointAchieved, desiredPoint) > distance)
                {
                    nearestPointAchieved[0] = currentPos[0];
                    nearestPointAchieved[1] = currentPos[1];
                    nearestPointAchieved[2] = currentPos[2];
                }
                Console.WriteLine($"Iteration number: {i + 1}   current position: ({currentPos[0]}, {currentPos[1]}, {currentPos[2]})   distance: {distance}");
                Console.WriteLine();
                if (currentPos[0] == 0 && currentPos[1] == 0 && currentPos[2] == 0)
                {
                    Console.WriteLine($"Nearest point achieved: ({nearestPointAchieved[0]}, {nearestPointAchieved[1]}, {nearestPointAchieved[2]})");
                    break;
                }
                else if (i == 999)
                    Console.WriteLine($"Nearest point achieved: ({nearestPointAchieved[0]}, {nearestPointAchieved[1]}, {nearestPointAchieved[2]})");
            }

            // Generate Excel File with timestamp in the name
            GenerateExcelFile(generatedPoints, $"GeneratedPoints_{DateTime.Now:yyyyMMddHHmmss}.xlsx");
        }

        static void GenerateExcelFile(List<(double[], double)> points, string fileName)
        {
            FileInfo fileInfo = new FileInfo(fileName);
            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Points");

                // Adding headers
                worksheet.Cells[1, 1].Value = "X";
                worksheet.Cells[1, 2].Value = "Y";
                worksheet.Cells[1, 3].Value = "Z";
                worksheet.Cells[1, 4].Value = "Distance";

                // Adding data
                for (int i = 0; i < points.Count; i++)
                {
                    worksheet.Cells[i + 2, 1].Value = points[i].Item1[0];
                    worksheet.Cells[i + 2, 2].Value = points[i].Item1[1];
                    worksheet.Cells[i + 2, 3].Value = points[i].Item1[2];
                    worksheet.Cells[i + 2, 4].Value = points[i].Item2;
                }

                package.Save();
            }
        }
    }
}
