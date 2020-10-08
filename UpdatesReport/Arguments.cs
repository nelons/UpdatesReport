using System;
using System.Collections.Specialized;
using System.Text.RegularExpressions;

namespace UpdatesReport
{
    class Arguments
    {
        private StringDictionary parameters;

        public Arguments(string[] Args)
        {
            parameters = new StringDictionary();

            // Define -, --, /, : and = as valid delimiters.  Ignore : and = if enclosed in quotes.
            Regex validDelims = new Regex(@"^-{1,2}|^/|[^['""]?.*]=['""]?$|[^['""]?.*]:['""]?$", RegexOptions.IgnoreCase | RegexOptions.Compiled);

            // Define anything enclosed with double quotes as a match.  We'll use this to replace
            // the entire string with only the part that matches (everything but the quotes)
            Regex quotedString = new Regex(@"^['""]?(.*?)['""]?$", RegexOptions.IgnoreCase | RegexOptions.Compiled);

            string currentParam = null;
            string[] parts;

            foreach (string arg in Args)
            {
                // Apply validDelims to current arg to see how many significant characters were present
                // We're limiting to 3 to forcefully ignore any characters in the parameter VALUE
                // that would normally be used as a delimiter
                parts = validDelims.Split(arg, 3);

                switch (parts.Length)
                {
                    // no special characters present.  we assume this means that this part
                    // represents a value to the previously provided parameter.
                    // For example, if we have: "--MyTestArg myValue"
                    // currentParam would currently be set to "--MyTestArg"
                    // parts[0] would hold "myValue", to be assigned to MyTestArg
                    case 1:
                        if (currentParam != null)
                        {
                            if (!parameters.ContainsKey(currentParam))
                                parameters.Add(currentParam, quotedString.Replace(parts[0], "$1"));

                            currentParam = null;
                        }
                        break;

                    // One split ocurred, meaning we found a parameter delimiter
                    // at the start of arg, but nothing to denote a value.
                    // example: --MyParam
                    case 2:
                        // We already had a parameter with no value last time through the loop.
                        // That means we have no explicit value to give currentParam. We'll default it to "true"
                        if (currentParam != null && !parameters.ContainsKey(currentParam))
                            parameters.Add(currentParam, "true");

                        // Store our value-less param and grab the next arg to see if it has our value
                        // parts[0] only contains the opening delimiter -, --, or /, 
                        // so we go after parts[1] for the actual param name
                        currentParam = parts[1];
                        break;

                    // Two splits occurred.  We found a starting parameter delimiter,
                    // a parameter name, and another delimiter denoting a value for this parameter
                    // Example: --MyParam=MyValue   or   --MyParam:MyValue
                    case 3:
                        // We already had a parameter with no value last time through the loop.
                        // That means we have no explicit value to give currentParam. We'll default it to "true"
                        if (currentParam != null && !parameters.ContainsKey(currentParam))
                            parameters.Add(currentParam, "true");

                        // Store the good param name
                        currentParam = parts[1];

                        // Ignores parameters that have already been presented, not thrilled about this approach...
                        if (!parameters.ContainsKey(currentParam))
                            parameters.Add(currentParam, quotedString.Replace(parts[2], "$1"));

                        // Reset currentParam, we already have both parameter and value for this arg
                        currentParam = null;
                        break;
                }
            }

            // Final cleanup, we may still have a parameter at the end of the args string that didn't get a value
            if (currentParam != null)
            {
                if (!parameters.ContainsKey(currentParam))
                    parameters.Add(currentParam, "true");
            }

           // return parameters;
        }

        /*private StringDictionary Parameters;

        // Constructor
        public Arguments(string[] Args)
        {
            Parameters = new StringDictionary();
            Regex Spliter = new Regex(@"^-{1,2}|^/|=|:",
                RegexOptions.IgnoreCase|RegexOptions.Compiled);

            Regex Remover = new Regex(@"^['""]?(.*?)['""]?$",
                RegexOptions.IgnoreCase|RegexOptions.Compiled);

            string Parameter = null;
            string[] Parts;

            // Valid parameters forms:
            // {-,/,--}param{ ,=,:}((",')value(",'))
            // Examples: 
            // -param1 value1 --param2 /param3:"Test-:-work" 
            //   /param4=happy -param5 '--=nice=--'
            foreach(string Txt in Args)
            {
                // Look for new parameters (-,/ or --) and a
                // possible enclosed value (=,:)
                Parts = Spliter.Split(Txt,3);

                switch(Parts.Length){
                // Found a value (for the last parameter 
                // found (space separator))
                case 1:
                    if(Parameter != null)
                    {
                        if(!Parameters.ContainsKey(Parameter)) 
                        {
                            Parts[0] = 
                                Remover.Replace(Parts[0], "$1");

                            Parameters.Add(Parameter, Parts[0]);
                        }
                        Parameter=null;
                    }
                    // else Error: no parameter waiting for a value (skipped)

                    break;

                // Found just a parameter

                case 2:
                    // The last parameter is still waiting. 

                    // With no value, set it to true.

                    if(Parameter!=null)
                    {
                        if(!Parameters.ContainsKey(Parameter)) 
                            Parameters.Add(Parameter, "true");
                    }
                    Parameter=Parts[1];
                    break;

                // Parameter with enclosed value

                case 3:
                    // The last parameter is still waiting. 

                    // With no value, set it to true.

                    if(Parameter != null)
                    {
                        if(!Parameters.ContainsKey(Parameter)) 
                            Parameters.Add(Parameter, "true");
                    }

                    Parameter = Parts[1];

                    // Remove possible enclosing characters (",')

                    if(!Parameters.ContainsKey(Parameter))
                    {
                        Parts[2] = Remover.Replace(Parts[2], "$1");
                        Parameters.Add(Parameter, Parts[2]);
                    }

                    Parameter=null;
                    break;
                }
            }
            // In case a parameter is still waiting

            if(Parameter != null)
            {
                if(!Parameters.ContainsKey(Parameter)) 
                    Parameters.Add(Parameter, "true");
            }
        }*/

        // Retrieve a parameter value if it exists 

        // (overriding C# indexer property)

        public string this [string Param]
        {
            get
            {
                return(parameters[Param]);
            }
        }
    }
}
