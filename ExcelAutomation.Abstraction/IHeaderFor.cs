﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAutomation.Abstraction
{
    public interface IHeaderFor<in TInput, out TOutput> 
    {
        TOutput Execute(TInput input);
    }
}