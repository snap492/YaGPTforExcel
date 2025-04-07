using RestSharp;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using System;
using Newtonsoft.Json;
using System.Windows.Forms;  // Для работы с JSON-ответом

namespace YaGPTforExcel.Services
{
    public class Yagpt4excelService
    {
        private const string API_URL = "https://llm.api.cloud.yandex.net/foundationModels/v1/completion";
        private string iamToken;
        private string folderId;
        private string token;

        public Yagpt4excelService(string token, string folderId)
        {
            this.token = token;
            this.folderId = folderId;
        }

        // Метод для получения нового IAM токена через OAuth
        public async Task<string> GetIamTokenAsync(string oauthToken)
        {
            System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;

            var client = new RestClient("https://iam.api.cloud.yandex.net/iam/v1/tokens");
            var request = new RestRequest("", Method.Post);

            var body = new
            {
                yandexPassportOauthToken = oauthToken
            };

            request.AddJsonBody(body);

            var response = await client.ExecuteAsync(request);

            if (response.IsSuccessful)
            {
                // Если запрос успешен, возвращаем токен из ответа
                var tokenData = JsonConvert.DeserializeObject<IamTokenResponse>(response.Content);
                return tokenData?.IamToken ?? "IAM токен не найден в ответе.";
            }
            else
            {
                // Если запрос не успешен, выводим ошибку               
                Console.WriteLine($"Ошибка: {response.StatusCode}");
                Console.WriteLine($"Тело ответа: {response.Content}");
                Console.WriteLine($"Сообщение: {response.ErrorMessage}");
                return $"Ошибка: {response.StatusDescription}";
            }
        }
        //Метод для отправки промта
        public async Task<string> GenerateText(string prompt)
        {
            System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;

            iamToken = await GetIamTokenAsync(token); // Получаем новый IAM токен

            // Создаем клиента для отправки запросов
            var client = new RestClient(API_URL);

            // Создаем запрос, указываем путь и метод POST
            var request = new RestRequest();
            request.Method = Method.Post;

            // Добавляем заголовки
            request.AddHeader("Authorization", $"Bearer {iamToken}");
            request.AddHeader("Content-Type", "application/json");

            // Формируем тело запроса
            var body = new
            {
                modelUri = $"gpt://{folderId}/yandexgpt",
                completionOptions = new { stream = false, temperature = 0.4, maxTokens = 2000 },
                messages = new[]
                {
                    new { role = "user", text = prompt }
                }
            };

            // Добавляем тело запроса в формате JSON
            request.AddJsonBody(body);

            // Отправляем запрос и получаем ответ
            var response = await client.ExecuteAsync(request);

            if (response.IsSuccessful && !string.IsNullOrEmpty(response.Content))
            {
                try
                {
                    dynamic json = Newtonsoft.Json.JsonConvert.DeserializeObject(response.Content);
                    return json?.result?.alternatives?[0]?.message?.text ?? "Ответ пуст";
                }
                catch
                {
                    return "Ошибка обработки ответа";
                }
            }
            else
            {
                return $"Ошибка запроса: {response.StatusCode}\n{response.Content}";
            }
        }
    }
    public class IamTokenResponse
    {
        [JsonProperty("iamToken")]
        public string IamToken { get; set; }

        [JsonProperty("expiresAt")]
        public string ExpiresAt { get; set; }
    }
}
