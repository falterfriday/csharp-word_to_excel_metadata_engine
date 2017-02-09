using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace w2e_conversion_test
{
    public class Parser
    {
        //DECLARING ALL THE NEEDED VARIABLES FOR SCOPE
        public List<Dictionary<string, string>> conversionList = new List<Dictionary<string, string>>();

        private string
            ceeNumber,
            title,
            typeOfQuestion,
            standardText,
            instructionsMarkup,
            lowestScoringReplica,
            comment,
            questionDescriptionMarkup,
            questionTemplate,
            responseLabel,
            responseSymbol,
            scoringText,
            isLastQuestion,
            nextQuestionText;

        int numberOfQuestions = 0;

        public void CheckText(string text, int columnNumber)
        {
            try
            {
                if (columnNumber == 1)
                {
                    if (text.StartsWith("CEE #:"))
                    {
                        Cee(text);
                    }
                    else if (text.StartsWith("STANDARD"))
                    {
                        Standard(text);
                    }
                    else if (text.StartsWith("Instructions:"))
                    {
                        Instructions(text);
                    }
                    else if (text.StartsWith("Lowest Scoring Replica:"))
                    {
                        LowestScoringReplica(text);
                    }
                    else if (text.Equals("Comment:"))
                    {
                        Comment(text);
                    }
                    else if (text.StartsWith("Q"))
                    {
                        numberOfQuestions++;
                    }
                }
                else if (columnNumber == 2)
                {
                    if (text.StartsWith("SCORE"))
                    {
                        Console.WriteLine("MADE IT TO THE SCORE!!");
                        NextQuestion(text);
                        Output();
                    }
                    else if (text.StartsWith("If") && text.Contains("Q"))
                    {
                        NextQuestion(text);
                        AddToArray();
                    }
                    else if (!(text.Equals("Question") || text.Equals("Response") || text.Equals("Scoring")))
                    {
                        Question(text);
                    }
                }
                else if (columnNumber == 3)
                {
                    if ((text.StartsWith("Y") || text.Contains("%") || text.Contains("#")))
                    {
                        Response(text);
                    }
                }
                else if (columnNumber == 4)
                {
                    Scoring(text);
                }
                
            }
            catch (Exception)
            {
                Console.WriteLine("Uh Oh... Something broke in the Parser.");
            }
        }

        ExcelWriter writer = new ExcelWriter();
        
//----------------------------WARNING TO YE: HERE BE DRAGONS---------------------------------
        
        
        

        private void Cee(string text)
        {
            //CEE QUESTION NUMBER IS GENERATED BUT ISN'T COMPLETE DUE TO UNKNOWN NUMBER OF QUESTIONS
            int colonIdx = text.IndexOf(":");
            int bracketIdx = text.IndexOf("]");

            //CEE NUMBER IS GENERATED
            ceeNumber = (text.Substring(colonIdx + 1, bracketIdx - colonIdx)).Trim();

            //TITLE IS GENERATED
            title = (text.Substring(bracketIdx + 1)).Trim();
        }

        private void Standard(string text)
        {
            int colonIdx = text.IndexOf(":");
            standardText = "<b>STANDARD:</b> " + (text.Substring(colonIdx + 2)).Trim();
        }
        
        private void Instructions(string text)
        {
            instructionsMarkup = text;
        }

        private void LowestScoringReplica(string text)
        {
            lowestScoringReplica = "Lowest Scoring Replica: <input type='text' id='lsr' />";
        }

        private void Comment(string text)
        {
            comment = "Comment: <input type='text' id='" + ceeNumber + "_COMM'";
        }
        
        private void Question(string text)
        {
            questionDescriptionMarkup += text;
        }

        private void Response(string text)
        {
            if (text.StartsWith("Y"))
            {
                questionTemplate = "YesNo";
            }
            else if (text.Contains("%") || text.Contains("#"))
            {
                questionTemplate = "input";
            }
        }

        private void Scoring(string text)
        {
            scoringText = text;
        }
        private void NextQuestion(string text)
        {
            nextQuestionText = text;
        }

        private void AddToArray()
        {
            Console.WriteLine("MADE IT TO THE OUTPUT METHOD");
            Console.WriteLine("number of questions = " + numberOfQuestions);

            conversionList.Add(new Dictionary<string, string>());
            conversionList[numberOfQuestions - 1].Add("ceeQuestionNumber", ceeNumber + "_Q" + numberOfQuestions);
            conversionList[numberOfQuestions - 1].Add("ceeNumber", ceeNumber);
            conversionList[numberOfQuestions - 1].Add("questionNumber", "Q" + numberOfQuestions);
            conversionList[numberOfQuestions - 1].Add("title", title);
            conversionList[numberOfQuestions - 1].Add("standardText", standardText);
            conversionList[numberOfQuestions - 1].Add("instructionsMarkup", instructionsMarkup);
            conversionList[numberOfQuestions - 1].Add("comment", comment);
            conversionList[numberOfQuestions - 1].Add("lowestScoringReplica", lowestScoringReplica);
            conversionList[numberOfQuestions - 1].Add("questionDescriptionMarkup", questionDescriptionMarkup);
            
            Console.WriteLine("MADE IT TO THE OUTPUT METHOD 2");
        }

        private void Output()
        {
            Writer(conversionList);
        }
        
        private void Writer(List<Dictionary<string, string>> conversionList)
        {
            Console.WriteLine("MADE IT TO THE WRITER METHOD");
            writer.WriteToExcel(conversionList);
        }
    }
}
