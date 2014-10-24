using System;
using System.Collections.Generic;
using System.Text;
using TrackingObject;

namespace TCL
{
    class Test
    {
        static void Main(string[] args)
        {
            TrackingCol TC = new TrackingCol();
            DateTime sdt = System.DateTime.Now;
            TC.Start = sdt.AddMonths(-2).ToString("dd/MM/yyyy");
            if (TC.Extract())
            {
                Console.WriteLine("Liste des " + TC.Count + " Evenements entre " + TC.Start + " et " + TC.End);
                Console.WriteLine();

                foreach (TrackingObj TO in TC)
                {
                    Console.WriteLine("Start Date : [" + TO.Start + "] / End Date [" + TO.End + "] - Subject [" + TO.Subject + "]");
                    foreach (string CAT in TO.Categories)
                    {
                        if ((CAT == "CAL:Avant Vente") || (CAT == "CAL:Avant Vente (OOF)") || (CAT == "CAL:Avant Vente (ConfCall)"))
                        {
                            if (TO.Subject.StartsWith("STU") != true)
                            {
                                TO.Subject = "STUB- " + TO.Subject;
//                                TO.Update();
                                Console.WriteLine("Trigger Update new subjet :" + TO.Subject);
                            }
                        }
                    }
                }
            }
        }
    }
}
