// Em MaterialPropertyUpdater.cs

using FormsAndWpfControls;
 // Garanta que este namespace esteja correto, se FunçõesExternas estiver nele.
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Xml;
using System.Xml.Linq;

namespace FormsAndWpfControls // Ou o namespace do seu projeto
{
    public static class MaterialPropertyUpdater
    {
        // ... (Seu método UpdateMaterialCustomProperties e ExemploAtualizacaoMaterial existentes) ...

        /// <summary>
        /// Obtém as propriedades customizadas de um material específico de um arquivo .sldmat.
        /// </summary>
        /// <param name="sldmatFilePath">Caminho completo para o arquivo .sldmat.</param>
        /// <param name="materialName">Nome do material a ser buscado.</param>
        /// <returns>
        /// Um dicionário onde a chave é o nome da propriedade e o valor é um dicionário
        /// com os atributos da propriedade (e.g., "value", "description", "units").
        /// Retorna null se o material não for encontrado ou se ocorrer um erro.
        /// Retorna um dicionário vazio se o material for encontrado, mas não tiver propriedades customizadas.
        /// </returns>
        public static Dictionary<string, Dictionary<string, string>> GetMaterialCustomProperties(string sldmatFilePath, string materialName)
        {
            if (!File.Exists(sldmatFilePath))
            {
                // Não mostra MessageBox aqui, pois pode ser chamado muitas vezes.
                // A UI que chama este método deve lidar com a ausência do arquivo.
                Console.WriteLine($"Erro (GetMaterialCustomProperties): O arquivo .sldmat '{Path.GetFileName(sldmatFilePath)}' não foi encontrado.");
                return null; // Indica falha na leitura do arquivo ou material
            }

            try
            {
                XDocument doc = XDocument.Load(sldmatFilePath);
                var searchResult = FuncoesExternas.BuscarMaterialRecursivo(doc.Root, materialName);

                if (searchResult.material == null)
                {
                    Console.WriteLine($"Aviso (GetMaterialCustomProperties): O material '{materialName}' não foi encontrado no arquivo '{Path.GetFileName(sldmatFilePath)}'.");
                    return null; // Indica que o material não foi encontrado.
                }

                XElement materialElement = searchResult.material;
                XElement customPropertiesElement = materialElement.Element("custom");

                var materialCustomProps = new Dictionary<string, Dictionary<string, string>>();

                if (customPropertiesElement != null && customPropertiesElement.Elements("prop").Any())
                {
                    foreach (var prop in customPropertiesElement.Elements("prop"))
                    {
                        string propName = prop.Attribute("name")?.Value ?? string.Empty;
                        if (!string.IsNullOrEmpty(propName))
                        {
                            var attributes = new Dictionary<string, string>();
                            foreach (XAttribute attr in prop.Attributes())
                            {
                                attributes[attr.Name.LocalName] = attr.Value;
                            }
                            materialCustomProps[propName] = attributes;
                        }
                    }
                }
                return materialCustomProps;
            }
            catch (XmlException xmlEx)
            {
                Console.WriteLine($"Erro XML (GetMaterialCustomProperties) ao processar '{Path.GetFileName(sldmatFilePath)}': {xmlEx.Message}");
                return null;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erro inesperado (GetMaterialCustomProperties) ao ler .sldmat: {ex.Message}");
                return null;
            }
        }

        public static bool UpdateMaterialCustomProperties(string sldmatFilePath, string materialName, Dictionary<string, Dictionary<string, string>> newProperties)
        {
            if (!File.Exists(sldmatFilePath))
            {
                Console.WriteLine($"Erro (UpdateMaterialCustomProperties): Arquivo '{Path.GetFileName(sldmatFilePath)}' não encontrado.");
                return false;
            }

            try
            {
                XDocument doc = XDocument.Load(sldmatFilePath);
                var searchResult = FuncoesExternas.BuscarMaterialRecursivo(doc.Root, materialName);

                if (searchResult.material == null)
                {
                    Console.WriteLine($"Erro (UpdateMaterialCustomProperties): Material '{materialName}' não encontrado em '{Path.GetFileName(sldmatFilePath)}'.");
                    return false;
                }

                XElement materialElement = searchResult.material;

                // Garante que exista o nó <custom>
                XElement customElement = materialElement.Element("custom");
                if (customElement == null)
                {
                    customElement = new XElement("custom");
                    materialElement.Add(customElement);
                }

                // Remove propriedades com mesmo nome (para evitar duplicatas)
                foreach (var newProp in newProperties)
                {
                    var existingProp = customElement.Elements("prop")
                        .FirstOrDefault(p => (string)p.Attribute("name") == newProp.Key);

                    existingProp?.Remove(); // Remove se já existir
                }

                // Adiciona todas as novas propriedades
                foreach (var newProp in newProperties)
                {
                    XElement propElement = new XElement("prop");

                    foreach (var attr in newProp.Value)
                    {
                        propElement.SetAttributeValue(attr.Key, attr.Value);
                    }

                    customElement.Add(propElement);
                }

                // Salva o arquivo com formatação
                doc.Save(sldmatFilePath);
                Console.WriteLine($"Propriedades do material '{materialName}' atualizadas com sucesso em '{Path.GetFileName(sldmatFilePath)}'.");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erro ao atualizar propriedades customizadas: {ex.Message}");
                return false;
            }
        }

    }
}