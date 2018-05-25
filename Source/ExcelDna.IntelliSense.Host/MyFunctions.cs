#if DEBUG
using System;
using System.ComponentModel;
using ExcelDna.Integration;
using ExcelDna.Logging;

namespace ExcelDna.CustomAddin
    {
        // These functions - are just here for testing...
        public class MyFunctions
        {
            [ExcelFunction(Description = "Returns the sum of two particular numbers that are given\r\n(As a test, of course)",
                           HelpTopic = "http://www.google.com")]
            public static double AddThem(
                [ExcelArgument(Name = "Augend", Description = "is the first number, to which will be added")] double v1,
                [ExcelArgument(Name = "Addend", Description = "is the second number that will be added")]     double v2)
            {
                return v1 + v2;
            }

            [ExcelFunction(Description = "--------------------",
                           HelpTopic = "MyFile.chm!100")]
            public static double AdxThem(
                [ExcelArgument(Name = "[tag]", Description = "is the first number, to which will be added")] double v1,
                [ExcelArgument(Name = "[Addend]", Description = "is the second number that will be added - and is described with a very long description that goes on and on in dangar of talking about foxes and dogs")]     double v2)
            {
                return v1 + v2;
            }

            [Description("Test function for the amazing Excel-DNA IntelliSense feature")]
            public static string jDummyFunc()
            {
                return "Howzit !";
            }

            [ExcelFunction(Name ="a.test?d_.3")]
            public static object AnotherFunction( [Description("In and out")] object inout)
            {
                return inout;
            }

            [ExcelFunction(Name ="TestArgs")]
            public static string TestArgs([ExcelArgument(Name = "First", Description = "First Arg")] object first,
                [ExcelArgument(Name = "Second", Description = "Second Arg;ArgList - hundred,thousand,million,billion,trillion")]
                string second)
        {
            return first.ToString() + " " + second.ToString();
        }

            [ExcelFunction(Name ="A.Non.Descript.Function")]
            public static object ANonDescriptFunction(object inout)
            {
                return inout;
            }

            [ExcelFunction(Name ="A.Descript.Function", Description = "Has a description")]
            public static object ADescriptFunction(object inout)
            {
                return inout;
            }

            [ExcelFunction(Description = "Has many arguments")]
            public static object AManyArgFunction(
                [ExcelArgument(Name = "Argument1", Description = "is the first argument")] double arg1,
                [ExcelArgument(Name = "Argument2", Description = "is another argument")] double arg2,
                [ExcelArgument(Name = "Argument3", Description = "is another argument")] double arg3,
                [ExcelArgument(Name = "Argument4", Description = "is another argument")] double arg4,
                [ExcelArgument(Name = "Argument5", Description = "is another argument")] double arg5,
                [ExcelArgument(Name = "Argument6", Description = "is another argument")] double arg6,
                [ExcelArgument(Name = "Argument7", Description = "is another argument")] double arg7,
                [ExcelArgument(Name = "Argument8", Description = "is another argument")] double arg8,
                [ExcelArgument(Name = "Argument9", Description = "is another argument")] double arg9,
                [ExcelArgument(Name = "Argument10", Description = "is another argument")] double arg10,
                [ExcelArgument(Name = "Argument11", Description = "is another argument")] double arg11,
                [ExcelArgument(Name = "Argument12", Description = "is another argument")] double arg12,
                [ExcelArgument(Name = "Argument13", Description = "is another argument")] double arg13,
                [ExcelArgument(Name = "Argument14", Description = "is another argument")] double arg14,
                [ExcelArgument(Name = "Argument15", Description = "is another argument")] double arg15,
                [ExcelArgument(Name = "Argument16", Description = "is another argument")] double arg16,
                [ExcelArgument(Name = "Argument18", Description = "is another argument")] double arg18,
                [ExcelArgument(Name = "Argument19", Description = "is another argument")] double arg19,
                [ExcelArgument(Name = "Argument20", Description = "is another argument")] double arg20
                )
            {
                return arg1;
            }

            [ExcelFunction(Description = @"Lorem ipsum dolor sit amet, consectetur adipiscing elit. Quisque vitae tortor faucibus nisl consectetur varius eget at leo. Praesent porta semper venenatis. Aliquam erat volutpat. Quisque erat urna, pharetra eu justo luctus, aliquam efficitur massa. Duis at vulputate nibh. Sed auctor leo non felis convallis varius. Vestibulum nec convallis nisl. Aenean quam elit, pulvinar at lectus nec, venenatis sollicitudin risus. Donec porta sapien sem, vitae tempus purus porttitor nec. Nunc orci lorem, eleifend quis massa ac, pellentesque accumsan tellus. Vestibulum consectetur cursus enim.

Maecenas vitae feugiat turpis.Proin at consequat enim.Etiam quis leo non tortor blandit fermentum quis vel mauris.Quisque maximus eros non sodales semper.Proin bibendum porttitor lobortis.Donec dictum cursus ante, pharetra luctus neque facilisis a.Aliquam quis ultrices enim, in molestie eros.

In placerat, arcu et ultricies ornare, orci nunc tristique justo, nec maximus odio mauris vitae enim.Proin fermentum dignissim mi, sed elementum enim aliquet id.Cras malesuada dignissim dui, ut euismod erat condimentum non.Quisque quis metus rhoncus, varius turpis vitae, suscipit libero.Phasellus felis tellus, euismod sed enim vestibulum, viverra sodales leo.Donec eu cursus purus, ac malesuada lacus.Nulla at sem quis tellus ullamcorper mattis vitae sit amet tortor.Curabitur in mauris ornare, maximus metus eu, pharetra orci.Sed pharetra felis sit amet lectus placerat accumsan.Sed porttitor ligula ac augue imperdiet euismod.Donec leo massa, sodales eget auctor a, luctus non mi.Curabitur et tellus sem.

Aliquam varius mauris sit amet rutrum cursus.Morbi tempus dui odio, quis bibendum dolor malesuada a.Suspendisse sit amet sodales nunc, non luctus lorem.Proin a tempus felis, a molestie ex.Donec euismod tellus quis quam blandit, non vehicula felis luctus.Aliquam eleifend, libero et mollis vulputate, est erat tincidunt nisi, non cursus purus nisl at ipsum.Ut mattis turpis vel sapien imperdiet, vitae dapibus lectus malesuada.Nullam ut dolor placerat, volutpat ex in, porta eros.Suspendisse potenti.Duis non dignissim urna.Morbi ipsum erat, convallis non ullamcorper et, tempus et magna.Morbi euismod mi hendrerit, faucibus arcu ut, lobortis eros.Maecenas a tellus nec massa viverra malesuada.Nam eu efficitur erat, nec tincidunt erat.

Ut rutrum sapien efficitur tellus auctor consectetur.Mauris eu mattis dui.Suspendisse id vestibulum est.Quisque posuere sem neque.Suspendisse quis lorem rutrum, commodo massa sed, mollis est.Maecenas accumsan mauris nec ligula lacinia dapibus.Sed efficitur, dolor eget varius vestibulum, ex libero ornare ligula, eu pharetra dolor lorem ut mauris.Duis metus ante, dapibus a mauris vitae, blandit gravida nulla.Aliquam vel mattis libero.Ut eu dignissim metus."
)]
            public static object VeryLongDescription()
        {
            return "True";
        }

        [ExcelCommand]
            public static void dnaLogDisplayShow()
            {
                LogDisplay.Show();
            }
        }
    }
#endif
