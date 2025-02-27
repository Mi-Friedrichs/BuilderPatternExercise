﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Office2021.Word;
using DocumentFormat.OpenXml.Office.Word;

namespace Exercise.WordBuilder.Models
{
    internal class DocumentData
    {
        private List<string> _TableOfContents;
        public List<string> TableOfContents
        {
            get
            {
                if (_TableOfContents == null)
                {
                    _TableOfContents = new List<string>();
                }

                return _TableOfContents;
            }
            set { _TableOfContents = value; }
        }

        public string Title { get; set; } = "Neues Dokument";

        /// <summary>
        /// Inhaltsverzeichnis generieren und ausgeben
        /// </summary>
        public bool IncludeTableOfContents { get; set; }

        /// <summary>
        /// Kapitelnummerierung
        /// </summary>
        public bool ChapterNumbering { get; set; }

        /// <summary>
        /// Vorwort
        /// </summary>
        public string Preface { get; set; } = "";

        /// <summary>
        /// Kapitel des Buches
        /// </summary>
        public List<Chapter> Chapters = new List<Chapter>();
    }
}
