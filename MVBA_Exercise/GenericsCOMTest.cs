///Code property of J Voss
///2022
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace MVBA_Exercise
{

    /// <summary>
    /// The purpose of this test class is a response to a job interview request with the MVBA Law Firm.
    /// Problem statement:
    /// Write a C# unit test to verify that all public generic classes in System.Collections.Generic that implement IEnumerable have a ComVisible attribute.
    /// The failure message should contain the names of all classes that do not meet the requirement, if any.
    /// 
    /// Two tests were written, one that employs LINQ and one that does not.  To account for shop preferences.
    /// Time taken was roughly 50 minutes for prettification  :-)
    /// </summary>
    [TestClass]
    public class GenericsCOMTest
    {
        /// <summary>
        /// Test ensures all public classes in the System.Collections.Generic namespace that implement
        /// <see cref="IEnumerable{T}<>"/> also implement the <see cref="ComVisibleAttribute"/>
        /// Test employs LINQ during execution which often increases readability 
        /// </summary>
        [TestMethod]
        public void TestEnsureIEnumerablesImplementCOMVisible_LINQ()
        {
            /// store the types we will be testing, make one call each to typeof()
            Type comVisibleAttrType = typeof(ComVisibleAttribute);
            Type ienumerableType = typeof(IEnumerable<>);

            /// create string variable for the namespace name
            string namespaceOfConcern = "System.Collections.Generic";

            /// create a list that will hold any and all {namespace}.{typename} for all types
            /// that match our filter criteria that do not implement ComVisible
            List<string> offenders = new List<string>();

            /// get all assemblies in our current AppDomain, then reflect through each for our test
            var assemblies = AppDomain.CurrentDomain.GetAssemblies();
            foreach (var assembly in assemblies)
            {
                /// Get a list of types that are in the correct namespace, that are classes and are publiic
                IEnumerable<Type> applicableTypes = assembly.GetTypes().Where(t => t.Namespace == namespaceOfConcern && t.IsPublic && t.IsClass);

                foreach (var type in applicableTypes)
                {
                    ///IEnumerables will implement Generic interface IEnumerable<>
                    if (type.GetInterfaces().Any(i => i.IsGenericType && i.GetGenericTypeDefinition() == ienumerableType))
                    {
                        ///COM Visible attribute should be in the custom attributes.  If it is not, we want to record it
                        if (type.GetCustomAttributes(comVisibleAttrType, true).Any() == false)
                        {
                            offenders.Add(string.Format("{0}.{1}", type.Namespace, type.Name));
                        }
                    }
                }
            }

            // our string builder will be used to create the fail message
            StringBuilder stringBuilder = new StringBuilder();

            if (offenders.Any())
            {
                stringBuilder.AppendLine("The following public classes were found NOT IMPLEMENTING the ComVisible attribute");
                foreach (var offender in offenders)
                {
                    stringBuilder.AppendLine(string.Format(" - {0}", offender));
                }
            }

            // our test fails if there are offenders.
            Assert.IsTrue(offenders.Count == 0, stringBuilder.ToString());


        }

        /// <summary>
        /// Test ensures all public classes in the System.Collections.Generic namespace that implement
        /// <see cref="IEnumerable{T}<>"/> also implement the <see cref="ComVisibleAttribute"/>
        /// Test does not employ LINQ which reduces readability (IMHO)
        /// </summary>
        [TestMethod]
        public void TestEnsureIEnumerablesImplementCOMVisible_LOOPS()
        {
            /// store the types we will be testing, make one call each to typeof()
            Type comVisibleAttrType = typeof(ComVisibleAttribute);
            Type ienumerableType = typeof(IEnumerable<>);

            /// create string variable for the namespace name
            string namespaceOfConcern = "System.Collections.Generic";

            /// create a list that will hold any and all {namespace}.{typename} for all types
            /// that match our filter criteria that do not implement ComVisible
            List<string> offenders = new List<string>();

            /// get all assemblies in our current AppDomain, then reflect through each for our test
            var assemblies = AppDomain.CurrentDomain.GetAssemblies();
            foreach (var assembly in assemblies)
            {
                /// Get all types for the assembly and loop through them looking for criteria
                IEnumerable<Type> applicableTypes = assembly.GetTypes();
                foreach (var applicableType in applicableTypes)
                {
                    /// first level test (namespace, class and public
                    if (applicableType.Namespace == namespaceOfConcern && applicableType.IsPublic && applicableType.IsClass)
                    {
                        /// get the interfaces implemented for the type and test
                        var implementedInterfaces = applicableType.GetInterfaces();
                        foreach (var implementedInterface in implementedInterfaces)
                        {
                            if (implementedInterface.IsGenericType && implementedInterface.GetGenericTypeDefinition() == ienumerableType)
                            {
                                ///COM Visible attribute should be in the custom attributes.  If it is not, we want to record it
                                if (applicableType.GetCustomAttributes(comVisibleAttrType, true).Any() == false)
                                {
                                    offenders.Add(string.Format("{0}.{1}", applicableType.Namespace, applicableType.Name));
                                    break;
                                }
                            }
                        }
                    }
                }
            }

            // our string builder will be used to create the fail message
            StringBuilder stringBuilder = new StringBuilder();

            if (offenders.Any())
            {
                stringBuilder.AppendLine("The following public classes were found NOT IMPLEMENTING the ComVisible attribute");
                foreach (var offender in offenders)
                {
                    stringBuilder.AppendLine(string.Format(" - {0}", offender));
                }
            }

            // our test fails if there are offenders.
            Assert.IsTrue(offenders.Count == 0, stringBuilder.ToString());


        }


    }
}
