using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using Excel = Microsoft.Office.Interop.Excel;

namespace WpfApp1.Pages
{
    /// <summary>
    /// Логика взаимодействия для Page2.xaml
    /// </summary>
    public partial class Page2 : Page
    {
        private int mode;
        public Page2(int mode)
        {
            InitializeComponent();
            this.mode = mode;
            ShowSelectedQueuePanel();
        }

        #region Обработчики событий



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
            CountThread.Text = null;
        }

        #endregion

        #region Чтение входных данных

        /// <summary>Читает количество каналов n с проверкой.</summary>
        private bool TryParseChannels(out int n)
        {
            n = 0;
            if (string.IsNullOrWhiteSpace(CountThread.Text))
            {
                ShowWarning("Заполните все обязательные поля (n, λ, μ / t_обс).");
                return false;
            }
            n = int.Parse(CountThread.Text);
            if (n <= 0)
            {
                ShowWarning("Количество каналов должно быть больше 0.");
                return false;
            }
            return true;
        }

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

        #region Вспомогательные вычисления

        private static double Factorial(int n)
        {
            if (n <= 1) return 1;
            double result = 1;
            for (int i = 2; i <= n; i++) result *= i;
            return result;
        }

        /// <summary>Сумма r^k / k! при k = 0..n.</summary>
        private static double SumPowFactorial(double r, int n)
        {
            double sum = 0;
            for (int k = 0; k <= n; k++)
            {
                sum += Math.Pow(r, k) / Factorial(k);
            }
            return sum;
        }

        #endregion

        #region Расчёты

        private void CalculateAnswer()
        {
            try
            {
                if (!TryParseChannels(out int n))
                {
                    return;
                }

                double lambda = ParseLambda();
                ParseServiceRate(out double mu, out double tObs);

                double r = lambda * tObs;   // интенсивность нагрузки
                double rho = r / n;         // коэффициент загрузки

                var sb = new StringBuilder();
                AppendHeader(sb, n, lambda, mu, tObs, r, rho);

                bool success = true;
                switch (mode)
                {
                    case 1:
                        AppendLossSystemResults(sb, n, lambda, r);
                        break;
                    case 2:
                        success = AppendLimitedQueueResults(sb, n, lambda, r);
                        break;
                    case 3:
                        AppendUnlimitedQueueResults(sb, n, lambda, r, rho, CountQ.Text);
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

        private static void AppendHeader(StringBuilder sb, int n, double lambda, double mu, double tObs, double r, double rho)
        {
            sb.AppendLine("Системы массового обслуживания (Многоканальные).");
            sb.AppendLine($"Количество каналов: n = {n}");
            sb.AppendLine($"Интенсивность потока заявок: λ = {lambda:F3} заявок/мин.");
            sb.AppendLine($"Интенсивность обслуживания: μ = {mu:F3} заявок/мин.");
            sb.AppendLine($"Среднее время обслуживания: t_обс = {tObs:F3} мин.");
            sb.AppendLine($"Интенсивность нагрузки: r = λ * t_обс = {r:F3}");
            sb.AppendLine($"Коэффициент загрузки: ρ = r/n = {rho:F3}\n");
        }

        /// <summary>Многоканальная СМО с отказами.</summary>
        private static void AppendLossSystemResults(StringBuilder sb, int n, double lambda, double r)
        {
            sb.AppendLine("--- Многоканальная СМО с отказами ---");

            double P0 = 1.0 / SumPowFactorial(r, n);
            sb.AppendLine($"Вероятность того, что все каналы свободны: P0 = {P0:F4}");

            double pOtk = (Math.Pow(r, n) / Factorial(n)) * P0;
            sb.AppendLine($"Вероятность отказа (все каналы заняты): P_отк = {pOtk:F4} ({pOtk * 100:F2}%)");

            double Q = 1 - pOtk;
            sb.AppendLine($"Относительная пропускная способность: Q = {Q:F4}");

            double A = lambda * Q;
            sb.AppendLine($"Абсолютная пропускная способность: A = {A:F3} заявок/мин.");

            double kBusy = r * Q;
            sb.AppendLine($"Среднее число занятых каналов: k = r * Q = {kBusy:F3}");

            double tSys = kBusy / lambda;
            sb.AppendLine($"Среднее время пребывания заявки в системе: t_сист = {tSys:F3} мин.");
        }

        /// <summary>Многоканальная СМО с ограниченной очередью. Возвращает false, если данные не введены.</summary>
        private bool AppendLimitedQueueResults(StringBuilder sb, int n, double lambda, double r)
        {
            if (!TryParseQueueLength(0, out int m))
            {
                return false;
            }

            sb.AppendLine($"--- Многоканальная СМО с ограниченной очередью (m = {m}) ---");

            double sumQueue = 0;
            for (int k = 1; k <= m; k++) sumQueue += Math.Pow(r / n, k);

            double P0 = 1.0 / (SumPowFactorial(r, n) + (Math.Pow(r, n) / Factorial(n)) * sumQueue);
            sb.AppendLine($"Вероятность того, что все каналы свободны: P0 = {P0:F4}");

            double pOtk = (Math.Pow(r, n + m) / (Factorial(n) * Math.Pow(n, m))) * P0;
            sb.AppendLine($"Вероятность отказа: P_отк = {pOtk:F4} ({pOtk * 100:F2}%)");

            double Q = 1 - pOtk;
            sb.AppendLine($"Относительная пропускная способность: Q = {Q:F4}");

            double A = lambda * Q;
            sb.AppendLine($"Абсолютная пропускная способность: A = {A:F3} заявок/мин.");

            double LOch = ComputeLimitedQueueLength(P0, n, r, m);
            sb.AppendLine($"Средняя длина очереди: L_оч = {LOch:F3}");

            double LSys = LOch + r * Q;
            sb.AppendLine($"Среднее число заявок в системе: L_сист = {LSys:F3}");

            sb.AppendLine($"Среднее время ожидания в очереди: T_оч = {LOch / lambda:F3} мин.");
            sb.AppendLine($"Среднее время в системе: T_сист = {LSys / lambda:F3} мин.");
            return true;
        }

        /// <summary>Средняя длина очереди для СМО с ограниченной очередью.</summary>
        private static double ComputeLimitedQueueLength(double P0, int n, double r, int m)
        {
            double L = 0;
            double coeff = (P0 * Math.Pow(r, n)) / Factorial(n);
            for (int k = 1; k <= m; k++) L += k * Math.Pow(r / n, k);
            return L * coeff;
        }

        /// <summary>Многоканальная СМО с неограниченной очередью.</summary>
        private static void AppendUnlimitedQueueResults(StringBuilder sb, int n, double lambda, double r, double rho, string CountQ)
        {
            sb.AppendLine("--- Многоканальная СМО с неограниченной очередью ---");

            if (rho >= 1)
            {
                sb.AppendLine($"ВНИМАНИЕ: Коэффициент загрузки ρ = {rho:F3} >= 1.");
                sb.AppendLine("Система не справляется с потоком. Очередь будет бесконечно расти.");
                return;
            }

            double P0 = 1.0 / (SumPowFactorial(r, n - 1) + (Math.Pow(r, n) / Factorial(n)) * (1.0 / (1 - rho)));
            sb.AppendLine($"Вероятность того, что все каналы свободны: P0 = {P0:F4}");

            sb.AppendLine("Вероятность отказа: P_отк = 0 (очередь не ограничена)");
            sb.AppendLine("Относительная пропускная способность: Q = 1");
            sb.AppendLine($"Абсолютная пропускная способность: A = λ = {lambda:F3} заявок/мин.");

            double PWait = (Math.Pow(r, n) / (Factorial(n) * (1 - rho))) * P0;
            sb.AppendLine($"Вероятность ожидания в очереди: P_ож = {PWait:F4}");

            double LOch = PWait * (rho / (1 - rho));
            sb.AppendLine($"Средняя длина очереди: L_оч = {LOch:F3}");

            double LSys = LOch + r;
            sb.AppendLine($"Среднее число заявок в системе: L_сист = {LSys:F3}");

            sb.AppendLine($"Среднее время ожидания в очереди: T_оч = {LOch / lambda:F3} мин.");
            sb.AppendLine($"Среднее время в системе: T_сист = {LSys / lambda:F3} мин.");

            if (!string.IsNullOrWhiteSpace(CountQ))
            {
                int m = int.Parse(CountQ);
                double Pm = PWait * Math.Pow(rho, m) * (1 - rho);
                sb.AppendLine($"Вероятность того, что в системе ровно {n + m} заявок (n каналов + {m} в очереди): {Pm:F4}");
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
                worksheet.Name = "Многоканальная СМО";

                string[] lines = TextAnswer.Text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                FillWorksheet(worksheet, lines);
                worksheet.Columns.AutoFit();
            }
            catch (Exception ex)
            {
                ShowError($"Ошибка экспорта в Excel: {ex.Message}\nУбедитесь, что Excel установлен.");
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
