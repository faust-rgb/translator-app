using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace TranslatorApp.Services;

/// <summary>
/// 提供 API Key 的安全存储服务，使用 Windows DPAPI 加密。
/// </summary>
public sealed class SecureApiKeyStorage : ISecureApiKeyStorage
{
    private static readonly byte[] Entropy = Encoding.UTF8.GetBytes("TranslatorApp_v1_");
    private readonly string _storagePath;

    public SecureApiKeyStorage()
    {
        _storagePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "TranslatorApp",
            "secure-keys.dat");
    }

    /// <inheritdoc />
    public Task<string?> GetAsync(string providerType)
    {
        try
        {
            if (!File.Exists(_storagePath))
            {
                return Task.FromResult<string?>(null);
            }

            var encryptedData = File.ReadAllBytes(_storagePath);
            var decryptedData = ProtectedData.Unprotect(encryptedData, Entropy, DataProtectionScope.CurrentUser);
            var json = Encoding.UTF8.GetString(decryptedData);

            var keys = JsonSerializer.Deserialize<Dictionary<string, string>>(json) ?? new Dictionary<string, string>();
            return Task.FromResult(keys.TryGetValue(providerType, out var key) ? key : null);
        }
        catch (CryptographicException)
        {
            // 解密失败，可能是数据损坏或用户变更
            return Task.FromResult<string?>(null);
        }
        catch (Exception)
        {
            return Task.FromResult<string?>(null);
        }
    }

    /// <inheritdoc />
    public async Task SetAsync(string providerType, string apiKey)
    {
        Dictionary<string, string> keys;

        try
        {
            if (File.Exists(_storagePath))
            {
                var encryptedData = await File.ReadAllBytesAsync(_storagePath);
                var decryptedData = ProtectedData.Unprotect(encryptedData, Entropy, DataProtectionScope.CurrentUser);
                var json = Encoding.UTF8.GetString(decryptedData);
                keys = JsonSerializer.Deserialize<Dictionary<string, string>>(json) ?? new Dictionary<string, string>();
            }
            else
            {
                keys = new Dictionary<string, string>();
            }
        }
        catch
        {
            keys = new Dictionary<string, string>();
        }

        if (string.IsNullOrEmpty(apiKey))
        {
            keys.Remove(providerType);
        }
        else
        {
            keys[providerType] = apiKey;
        }

        var directory = Path.GetDirectoryName(_storagePath)!;
        Directory.CreateDirectory(directory);

        var newJson = JsonSerializer.Serialize(keys);
        var newData = Encoding.UTF8.GetBytes(newJson);
        var encrypted = ProtectedData.Protect(newData, Entropy, DataProtectionScope.CurrentUser);

        await File.WriteAllBytesAsync(_storagePath, encrypted);
    }

    /// <inheritdoc />
    public Task RemoveAsync(string providerType)
    {
        return SetAsync(providerType, string.Empty);
    }

    /// <inheritdoc />
    public Task ClearAsync()
    {
        if (File.Exists(_storagePath))
        {
            File.Delete(_storagePath);
        }

        return Task.CompletedTask;
    }
}

/// <summary>
/// API Key 安全存储接口。
/// </summary>
public interface ISecureApiKeyStorage
{
    /// <summary>
    /// 获取指定提供商的 API Key。
    /// </summary>
    Task<string?> GetAsync(string providerType);

    /// <summary>
    /// 设置指定提供商的 API Key。
    /// </summary>
    Task SetAsync(string providerType, string apiKey);

    /// <summary>
    /// 移除指定提供商的 API Key。
    /// </summary>
    Task RemoveAsync(string providerType);

    /// <summary>
    /// 清除所有存储的 API Key。
    /// </summary>
    Task ClearAsync();
}
