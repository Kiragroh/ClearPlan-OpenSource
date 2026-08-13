using System;
using System.Security.Cryptography;
using System.Text;
using VMS.TPS.Common.Model.API;

namespace ClearPlan.Helpers
{
    public static class PrivacyHelper
    {
        public static string FormatUserForLog(User user, ClearPlanSettings settings)
        {
            string rawUser = BuildUserLabel(user);
            return settings.Privacy.LogUserIdentifiers
                ? MakeDelimitedSafe(rawUser)
                : BuildPseudonym("user", rawUser, settings);
        }

        public static string FormatUserForDisplay(User user)
        {
            return MakeDelimitedSafe(BuildUserLabel(user));
        }

        public static string FormatPatientIdForLog(Patient patient, ClearPlanSettings settings)
        {
            string rawPatientId = patient == null ? string.Empty : patient.Id;
            return settings.Privacy.LogPatientIdentifiers
                ? MakeDelimitedSafe(rawPatientId)
                : BuildPseudonym("patient", rawPatientId, settings);
        }

        public static string FormatPatientNameForLog(Patient patient, ClearPlanSettings settings)
        {
            if (settings.Privacy.LogPatientIdentifiers)
            {
                return MakeDelimitedSafe(patient == null ? string.Empty : patient.Name);
            }

            return FormatPatientIdForLog(patient, settings);
        }

        public static string FormatPatientForFilename(Patient patient, ClearPlanSettings settings)
        {
            if (settings.Privacy.UsePatientIdentifiersInFilenames)
            {
                string patientId = patient == null ? string.Empty : patient.Id;
                string lastName = patient == null ? string.Empty : patient.LastName;
                return MakeDelimitedSafe(patientId + "_" + lastName);
            }

            return FormatPatientIdForLog(patient, settings);
        }

        public static string MakeDelimitedSafe(string value)
        {
            return string.IsNullOrWhiteSpace(value)
                ? "unknown"
                : value.Replace(",", "_").Replace(";", "_").Replace("\\", "_").Trim();
        }

        private static string BuildUserLabel(User user)
        {
            if (user == null)
            {
                return string.Empty;
            }

            string userId = string.IsNullOrWhiteSpace(user.Id) ? "unknown" : RemoveDomainPrefix(user.Id);
            string userName = string.IsNullOrWhiteSpace(user.Name) ? "unknown" : RemoveDomainPrefix(user.Name);
            return userId + "-" + userName;
        }

        private static string RemoveDomainPrefix(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            string[] parts = value.Split(new[] { '\\' }, StringSplitOptions.RemoveEmptyEntries);
            return parts.Length == 0 ? value : parts[parts.Length - 1];
        }

        private static string BuildPseudonym(string prefix, string value, ClearPlanSettings settings)
        {
            string normalized = string.IsNullOrWhiteSpace(value) ? "unknown" : value.Trim().ToUpperInvariant();
            string salt = settings.Privacy.HashSalt;
            if (string.IsNullOrWhiteSpace(salt))
            {
                salt = Environment.MachineName;
            }

            using (var sha256 = SHA256.Create())
            {
                byte[] bytes = Encoding.UTF8.GetBytes(salt + "|" + normalized);
                byte[] hash = sha256.ComputeHash(bytes);
                var builder = new StringBuilder();
                for (int i = 0; i < 6; i++)
                {
                    builder.Append(hash[i].ToString("x2"));
                }

                return prefix + "-" + builder;
            }
        }
    }
}
