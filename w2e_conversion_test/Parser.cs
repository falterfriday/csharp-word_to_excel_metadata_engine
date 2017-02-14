using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace w2e_conversion_test
{
    public class Parser
    {

        //DECLARING ALL THE NEEDED VARIABLES FOR SCOPE

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
            nextQuestionText,
            quesScoreColorRule,
            nextQuesBehaviorRule,
            responseCalculationRule,
            responseType;

        int numberOfQuestions = 0,
            counter = 0;

        public void CheckText(List<Dictionary<string, string>> conversionList, string text, int columnNumber)
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
                        NextQuestion(text);
                        PushToList(conversionList);
                        numberOfQuestions = 0;
                        conversionList.Last()["isLastQuestion"] = "TRUE";
                    }
                    else if (text.StartsWith("If") && text.Contains("Q"))
                    {
                        NextQuestion(text);
                        PushToList(conversionList);
                    }
                    else if (!(text.Equals("Question")))
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
        
//----------------------------WARNING: HERE BE DRAGONS---------------------------------
        
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
            comment = ceeNumber + "_COMM";
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
                typeOfQuestion = "RadioButtons";
                responseLabel = "";
                responseSymbol = "";
            }
            else if (text.Contains("%") || text.Contains("#"))
            {
                questionTemplate = "Input";
                typeOfQuestion = "Number";
                if (text.Contains("#"))
                {
                    responseLabel = "# Ticked";
                    responseSymbol = "";
                }
                else
                {
                    responseLabel = "";
                    responseSymbol = "%";
                }
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

        private void PushToList(List<Dictionary<string, string>> conversionList)
        {
            //ADD A DICTIONARY TO THE LIST
            conversionList.Add(new Dictionary<string, string>());
            
            //POPULATES THE DICTIONARY W/ KEY:VAL PAIRS
            conversionList[counter].Add("ceeQuestionNumber", ceeNumber + "_Q" + numberOfQuestions);
            conversionList[counter].Add("ceeNumber", ceeNumber);
            conversionList[counter].Add("questionNumber", "Q" + numberOfQuestions);
            conversionList[counter].Add("typeOfQuestion", typeOfQuestion);
            conversionList[counter].Add("title", title);
            conversionList[counter].Add("standardText", standardText);
            conversionList[counter].Add("instructionsMarkup", instructionsMarkup);
            conversionList[counter].Add("lowestScoringReplica", lowestScoringReplica);
            conversionList[counter].Add("comment", comment);
            conversionList[counter].Add("questionDescriptionMarkup", questionDescriptionMarkup);
            conversionList[counter].Add("questionTemplate", questionTemplate);
            conversionList[counter].Add("responseLabel", responseLabel);
            conversionList[counter].Add("responseSymbol", responseSymbol);
            conversionList[counter].Add("scoringText", scoringText);
            conversionList[counter].Add("isLastQuestion", "FALSE");
            conversionList[counter].Add("nextQuestionText", nextQuestionText);
            conversionList[counter].Add("quesScoreColorRule", quesScoreColorRule);
            conversionList[counter].Add("nextQuesBehaviorRule", nextQuesBehaviorRule);
            conversionList[counter].Add("responseCalculationRule", responseCalculationRule);
            conversionList[counter].Add("responseType", responseType);

            questionDescriptionMarkup = String.Empty;
            counter++;
        }

    }
}