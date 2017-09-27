using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LINQtoCSV.Excel.Tests
{
    public interface IAssertable<T>
    {
        void AssertEqual(T other);
    }
}
