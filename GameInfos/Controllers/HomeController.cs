﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using GameInfos.Models;
using HtmlAgilityPack;

namespace GameInfos.Controllers
{
    public class HomeController : Controller
    {
        public Task<IActionResult> Index()
        {
            return ExportAsExcelFile();
        }

        public async Task<IActionResult> ExportAsExcelFile()
        {
            var url =
                "https://play.google.com/store/apps/collection/cluster?clp=0g4gCh4KGHRvcHNlbGxpbmdfbmV3X2ZyZWVfR0FNRRAHGAM%3D:S:ANO1ljIxjbU&gsr=CiPSDiAKHgoYdG9wc2VsbGluZ19uZXdfZnJlZV9HQU1FEAcYAw%3D%3D:S:ANO1ljJL0LQ";


            var httpClient = new HttpClient();
            var html = await httpClient.GetStringAsync(url);

            var htmlDocument = new HtmlDocument();
            htmlDocument.LoadHtml(html);

            var gamesHtml = htmlDocument.DocumentNode.Descendants("div").Where(node => node.GetAttributeValue("class", "").Equals("ZmHEEd ")).ToList();
            var gameList = gamesHtml[0].Descendants("div").Where(x => x.GetAttributeValue("class", "").Contains("ImZGtf ")).ToList();
            var list = new List<string>();
            foreach (var game in gameList)
            {
                for (int i = 0; i < gameList.Count; i++)
                {
                    var gameLinkList = gameList[i].Descendants("a").Where(x => x.GetAttributeValue("href", "").Contains("/store/apps/details?id"));
                    foreach (var innerGameLink in gameLinkList)
                    {
                        if (!list.Contains(innerGameLink.GetAttributeValue("href", "")))
                        {
                            list.Add(innerGameLink.GetAttributeValue("href", ""));
                        }
                    }
                }
            }

            var urlList = new List<string>();

            foreach (var gameUrl in list)
            {
                var insert = gameUrl.Insert(0, "https://play.google.com");
                urlList.Add(insert);
            }


            var returnModel = new GameModel();


            foreach (var urlOfGame in urlList)
            {
                var httpClientFor = new HttpClient();
                var htmlFor = httpClientFor.GetStringAsync(urlOfGame);

                var htmlDocumentFor = new HtmlDocument();
                htmlDocumentFor.LoadHtml(await htmlFor);
                //Name of the Game
                var gameTitle = htmlDocumentFor.DocumentNode.SelectSingleNode(".//h1[@class='AHFaub']").InnerText;
                // Console.WriteLine(gameTitle); // test

                //Install count
                var outerInstallCount = htmlDocumentFor.DocumentNode.Descendants("div").Where(node => node.GetAttributeValue("class", "").Equals("hAyfc")).ToList();
                var installCount = outerInstallCount[2].SelectSingleNode(".//span").InnerText;
                // Console.WriteLine(installCount); //test

                //Company Name
                var companyName = htmlDocumentFor.DocumentNode.SelectSingleNode(".//span[@class='T32cc UAO9ie']").InnerText;
                // Console.WriteLine(companyName); // test

                //Company Email
                var companyEmail = htmlDocumentFor.DocumentNode.SelectSingleNode(".//a[@class='hrTbp euBY6b']").InnerText;
                // Console.WriteLine(companyEmail); // test

                //Country
                var divCountry = htmlDocumentFor.DocumentNode.Descendants("div").Where(node => node.GetAttributeValue("class", "").Equals("hAyfc")).ToList();
                var outerCountry = divCountry.LastOrDefault()?.Descendants("div").Where(node => node.GetAttributeValue("class", "").Equals("IQ1z0d")).ToList();
                var country = outerCountry[0].Descendants("div").LastOrDefault()?.InnerText;
                // Console.WriteLine(country); //test


                returnModel.GameName.Add(gameTitle);
                returnModel.InstallCount.Add(installCount);
                returnModel.CompanyName.Add(companyName);
                returnModel.Email.Add(companyEmail);
                if (country != null && country.Equals("Privacy Policy"))
                {
                    returnModel.Country.Add("Not mentioned");
                }
                else
                {
                      returnModel.Country.Add(country);
                }
              
               
            }


            List<GameModel> gameInfos = new List<GameModel>();
            gameInfos.AddRange(new[] {returnModel});


            var builder = new StringBuilder();
            builder.Append("Game Name, Install Count, Company Name, Email, Address");
            foreach (var game in gameInfos)
            {
                builder.AppendLine($"{game.GameName}, {game.InstallCount}, {game.CompanyName}, {game.Email}, {game.Country}");
            }

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("GameInfos");
                var currentRow = 1;
                worksheet.Cell(currentRow, 1).Value = "Game Name";
                worksheet.Cell(currentRow, 2).Value = "Install Count";
                worksheet.Cell(currentRow, 3).Value = "Company Name";
                worksheet.Cell(currentRow, 4).Value = "Email";
                worksheet.Cell(currentRow, 5).Value = "Address";

                foreach (var game in gameInfos)
                {
                    currentRow++;
                    worksheet.Cell(currentRow, 1).Value = game.GameName;
                    worksheet.Cell(currentRow, 2).Value = game.InstallCount;
                    worksheet.Cell(currentRow, 3).Value = game.CompanyName;
                    worksheet.Cell(currentRow, 4).Value = game.Email;
                    worksheet.Cell(currentRow, 5).Value = game.Country;
                }

                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var content = stream.ToArray();
                    return File(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "GameInfos.xlsx");
                }
            }
        }
    }
}