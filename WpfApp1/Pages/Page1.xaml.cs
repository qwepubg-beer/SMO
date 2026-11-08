using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using Excel = Microsoft.Office.Interop.Excel;

namespace WpfApp1.Pages
{
    /// <summary>
    /// Логика взаимодействия для Page1.xaml
    /// </summary>
    public partial class Page1 : Page
    {
        private const double Epsilon = 1e-4;
        private int mode;
        public Page1(int mode)
        {
            InitializeComponent();
            this.mode = mode;
            ShowSelectedQueuePanel();
        }

        #region Обработчики событий


        private void ChooseTimeIPZ_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
        }

        private void Answer_Click(object sender, RoutedEventArgs e)
        {
            CalculateAnswer();
        }

        private void Exsel_Click(object sender, RoutedEventArgs e)
        {
            ExportToExcel();
        }

        private void Clear_Click(object sender, RoutedEventArgs e)
        {
            ClearInputs();
        }

        #endregion

        #region Логика интерфейса

        private void ShowSelectedQueuePanel()
        {
            if (QueueLimitation == null || QueueWithoutlimitation == null) return;

            QueueLimitation.Visibility = Visibility.Collapsed;
            QueueWithoutlimitation.Visibility = Visibility.Collapsed;

            switch (mode)
            {
                case 2:
                    QueueLimitation.Visibility = Visibility.Visible;
                    break;
                case 3:
                    //QueueWithoutlimitation.Visibility = Visibility.Visible;
                    break;
            }
        }

        private void ClearInputs()
        {
            TextAnswer.Text = null;
            CountQ.Text = null;
            LenghtQ.Text = null;
            IPO.Text = null;
            IPZ.Text = null;
        }

        #endregion

        #region Чтение входных данных

        /// <summary>Читает интенсивность потока заявок λ с учётом единиц измерения (часы -> минуты).</summary>
        private double ParseLambda()
        {
            double lambda = double.Parse(IPZ.Text);
            if (ChooseTimeIPZ.SelectedIndex == 1)
            {
                lambda /= 60.0; // заявки/час -> заявки/мин
            }
            return lambda;
        }

        /// <summary>
        /// Читает интенсивность обслуживания: пользователь вводит либо μ,
        /// либо среднее время обслуживания t_обс — в зависимости от ServiceType.
        /// </summary>
        private void ParseServiceRate(out double mu, out double tObs)
        {
            double inputValue = double.Parse(IPO.Text);

            if (ServiceType.SelectedIndex == 0) // введено μ
            {
                mu = inputValue;
                if (Tops.SelectedIndex == 1) mu /= 60.0; // заявок/час -> заявок/мин
                tObs = 1.0 / mu;
            }
            else // введено t_обс
            {
                tObs = inputValue;
                if (Tops.SelectedIndex == 1) tObs *= 60.0; // часы -> минуты
                mu = 1.0 / tObs;
            }
        }

        /// <summary>Читает длину ограниченной очереди m с проверкой.</summary>
        private bool TryParseQueueLength(int defaultValue, out int m)
        {
            m = defaultValue;
            if (string.IsNullOrWhiteSpace(LenghtQ.Text))
            {
                ShowWarning("Введите длину очереди m.");
                return false;
            }
            m = int.Parse(LenghtQ.Text);
            return true;
        }

        #endregion

        #region Расчёты

        private void CalculateAnswer()
        {
            try
            {
                if (string.IsNullOrWhiteSpace(IPZ.Text) || string.IsNullOrWhiteSpace(IPO.Text))
                {
                    ShowWarning("Заполните значения интенсивности потока заявок и обслуживания.");
                    return;
                }

                double lambda = ParseLambda();
                ParseServiceRate(out double mu, out double tObs);
                double rho = lambda * tObs;

                var sb = new StringBuilder();
                AppendHeader(sb, mu, tObs);

                bool success = true;
                switch (this.mode)
                {
                    case 1:
                        AppendLossSystemResults(sb, lambda, tObs, rho);
                        break;
                    case 2:
                        success = AppendLimitedQueueResults(sb, lambda, rho);
                        break;
                    case 3:
                        AppendUnlimitedQueueResults(sb, lambda, rho, CountQ.Text);
                        break;
                }

                if (success)
                {
                    TextAnswer.Text = sb.ToString();
                }
            }
            catch (Exception ex)
            {
                ShowError($"Ошибка при расчете: {ex.Message}");
            }
        }

        private static void AppendHeader(StringBuilder sb, double mu, double tObs)
        {
            sb.AppendLine("Системы массового обслуживания.");
            sb.AppendLine($"Интенсивность потока обслуживания: μ = {mu:F3} заявок/мин.");
            sb.AppendLine($"Среднее время обслуживания: t_обс = 1/μ = {tObs:F3} мин.");
        }

        /// <summary>СМО с отказами (одно место, очереди нет).</summary>
        private static void AppendLossSystemResults(StringBuilder sb, double lambda, double tObs, double rho)
        {
            sb.AppendLine($"\nИнтенсивность нагрузки.\nρ = λ * t_обс = {lambda:F3} * {tObs:F3} = {rho:F3}");

            double p0 = 1.0 / (1.0 + rho);
            sb.AppendLine($"Вероятность, что канал свободен (доля времени простоя канала).\np0 = 1 / (1 + ρ) = 1 / (1 + {rho:F3}) = {p0:F3}");

            double pOtk = 1.0 - p0;
            sb.AppendLine($"Доля заявок, получивших отказ (вероятность отказа).\np_отк = 1 - p0 = 1 - {p0:F3} = {pOtk:F3}");

            double Q = p0;
            sb.AppendLine($"Относительная пропускная способность.\nДоля обслуживаемых заявок, поступающих в единицу времени:\nQ = p0 = {Q:F3}");

            double A = Q * lambda;
            sb.AppendLine($"Абсолютная пропускная способность.\nA = Q * λ = {Q:F3} * {lambda:F3} = {A:F3} заявок/мин.");

            double tPr = pOtk * tObs;
            sb.AppendLine($"Среднее время простоя СМО (мин.).\ntпр = p_отк * t_обс = {pOtk:F3} * {tObs:F3} = {tPr:F3}");

            double LObs = rho * Q;
            sb.AppendLine($"Среднее число обслуживаемых заявок.\nLобс = ρ * Q = {rho:F3} * {Q:F3} = {LObs:F3} ед.");
            sb.AppendLine($"Число заявок, получивших отказ: λ * p_отк = {(lambda * pOtk):F2} заявок/мин.");
            sb.AppendLine($"Номинальная производительность СМО: μ = {1.0 / tObs:F2} заявок/мин.");
            sb.AppendLine($"Фактическая производительность СМО: {A:F3} / {1.0 / tObs:F2} = {(A / (1.0 / tObs)) * 100:F0}% от номинальной.");
        }

        /// <summary>Одноканальная СМО с ограниченной очередью. Возвращает false, если данные не введены.</summary>
        private bool AppendLimitedQueueResults(StringBuilder sb, double lambda, double rho)
        {
            if (!TryParseQueueLength(0, out int m))
            {
                return false;
            }

            sb.AppendLine($"\nИсчисляем показатели для одноканальной СМО с ограниченной очередью (m={m}):");
            sb.AppendLine($"Интенсивность нагрузки ρ = {rho:F3}");

            double p0;
            if (Math.Abs(rho - 1.0) < Epsilon)
            {
                p0 = 1.0 / (m + 2);
                sb.AppendLine($"Вероятность того, что канал свободен (ρ=1): p0 = 1 / (m + 2) = {p0:F3}");
            }
            else
            {
                p0 = (1 - rho) / (1 - Math.Pow(rho, m + 2));
                sb.AppendLine($"Вероятность того, что канал свободен: p0 = (1 - ρ) / (1 - ρ^(m+2)) = {p0:F3}");
            }

            double pOtk = p0 * Math.Pow(rho, m + 1);
            sb.AppendLine($"Вероятность отказа: P_отк = p0 * ρ^(m+1) = {pOtk:F3}");

            double Q = 1 - pOtk;
            sb.AppendLine($"Относительная пропускная способность: Q = 1 - P_отк = {Q:F3}");

            double A = lambda * Q;
            sb.AppendLine($"Абсолютная пропускная способность: A = λ * Q = {A:F3} заявок/мин.");

            double LOch = (rho * rho * (1 - (m + 1) * Math.Pow(rho, m) + m * Math.Pow(rho, m + 1)))
                          / ((1 - Math.Pow(rho, m + 2)) * (1 - rho));
            sb.AppendLine($"Среднее число заявок в очереди: L_оч = {LOch:F3}");

            double LSist = LOch + Q;
            sb.AppendLine($"Среднее число заявок в системе: L_сист = L_оч + Q = {LSist:F3}");

            sb.AppendLine($"Среднее время ожидания в очереди: T_оч = L_оч / λ = {LOch / lambda:F3} мин.");
            sb.AppendLine($"Среднее время пребывания в системе: T_сист = L_сист / λ = {LSist / lambda:F3} мин.");
            return true;
        }

        /// <summary>Одноканальная СМО с неограниченной очередью.</summary>
        private static void AppendUnlimitedQueueResults(StringBuilder sb, double lambda, double rho, string CountQ)
        {
            sb.AppendLine("\nИсчисляем показатели для одноканальной СМО с неограниченной очередью:");

            if (rho >= 1)
            {
                sb.AppendLine($"ВНИМАНИЕ: Интенсивность нагрузки ρ = {rho:F3} >= 1.");
                sb.AppendLine("Система не справляется с потоком заявок, очередь будет расти до бесконечности.");
                return;
            }

            double p0 = 1 - rho;
            sb.AppendLine($"Вероятность того, что канал свободен: p0 = 1 - ρ = {p0:F3}");
            sb.AppendLine("Вероятность отказа: P_отк = 0 (система без отказов)");
            sb.AppendLine("Относительная пропускная способность: Q = 1");
            sb.AppendLine($"Абсолютная пропускная способность: A = λ = {lambda:F3} заявок/мин.");

            double LOch = (rho * rho) / (1 - rho);
            sb.AppendLine($"Среднее число заявок в очереди: L_оч = ρ^2 / (1 - ρ) = {LOch:F3}");

            double LSist = rho / (1 - rho);
            sb.AppendLine($"Среднее число заявок в системе: L_сист = ρ / (1 - ρ) = {LSist:F3}");

            sb.AppendLine($"Среднее время ожидания в очереди: T_оч = L_оч / λ = {LOch / lambda:F3} мин.");
            sb.AppendLine($"Среднее время пребывания в системе: T_сист = L_сист / λ = {LSist / lambda:F3} мин.");

            if (!string.IsNullOrWhiteSpace(CountQ))
            {
                int m = int.Parse(CountQ);
                double Pm = (1 - rho) * Math.Pow(rho, m);
                sb.AppendLine($"Вероятность нахождения m={m} заявок в системе: P_m = (1-ρ)*ρ^m = {Pm:F4}");
            }

        }

        #endregion

        #region Экспорт в Excel

        private void ExportToExcel()
        {
            if (string.IsNullOrWhiteSpace(TextAnswer.Text))
            {
                ShowInfo("Сначала выполните расчет (нажмите 'Решить').");
                return;
            }

            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;

            try
            {
                excelApp = new Excel.Application();
                excelApp.Visible = true;
                workbook = excelApp.Workbooks.Add(Type.Missing);
                worksheet = (Excel.Worksheet)workbook.ActiveSheet;
                worksheet.Name = "Результаты СМО";

                string[] lines = TextAnswer.Text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                FillWorksheet(worksheet, lines);
                worksheet.Columns.AutoFit();
            }
            catch (Exception ex)
            {
                ShowError($"Произошла ошибка при экспорте в Excel: {ex.Message}\nУбедитесь, что Excel установлен на компьютере.");
            }
            finally
            {
                ReleaseComObject(worksheet);
                ReleaseComObject(workbook);
                ReleaseComObject(excelApp);
            }
        }

        private static void FillWorksheet(Excel.Worksheet worksheet, string[] lines)
        {
            for (int i = 0; i < lines.Length; i++)
            {
                worksheet.Cells[i + 1, 1] = lines[i];
            }
        }

        private static void ReleaseComObject(object comObject)
        {
            if (comObject != null)
            {
                Marshal.ReleaseComObject(comObject);
            }
        }

        #endregion

        #region Сообщения пользователю

        private void ShowWarning(string message) =>
            MessageBox.Show(message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);

        private void ShowError(string message) =>
            MessageBox.Show(message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);

        private void ShowInfo(string message) =>
            MessageBox.Show(message, "Внимание", MessageBoxButton.OK, MessageBoxImage.Information);

        #endregion
    }
}
