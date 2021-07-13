using proyCentroImpresion.Models;
using System;
using System.Collections.Generic;
using System.Text;

namespace proyCentroImpresion.Data
{
    interface IReadFromExcel
    {
        List<WithheldExtract> GetWithheldExtractAsync();

        List<OutputType> GetOutpuTtypeAsync();
    }
}
