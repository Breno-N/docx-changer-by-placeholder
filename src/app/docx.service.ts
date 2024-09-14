import { Injectable } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { saveAs } from 'file-saver';
import PizZip from 'pizzip';
import * as Docxtemplater from 'docxtemplater';

@Injectable({
  providedIn: 'root',
})
export class DocxService {
  constructor(private http: HttpClient) {}

  /**
   * Processa o arquivo .docx e substitui os placeholders
   * @param fileUrl URL do arquivo .docx
   * @param placeholders Dados dos placeholders a serem substituídos
   * @param outputFileName Nome do novo arquivo gerado
   */
  public processDocxFile(
    fileUrl: string,
    placeholders: Record<string, string>,
    outputFileName: string
  ): void {
    this.fetchFile(fileUrl)
      .then((fileContent) => this.loadDocxFile(fileContent))
      .then((doc) => this.replacePlaceholders(doc, placeholders))
      .then((updatedDoc) => this.saveDocxFile(updatedDoc, outputFileName))
      .catch(this.handleError);
  }

  /**
   * Busca o arquivo .docx via HTTP
   * @param url URL do arquivo
   * @returns Promise com o conteúdo do arquivo como ArrayBuffer
   */
  private fetchFile(url: string): Promise<ArrayBuffer> {
    return this.http
      .get<ArrayBuffer>(url, { responseType: 'arraybuffer' as 'json' })
      .toPromise()
      .then((response) => {
        if (!response) throw new Error('Erro: arquivo não encontrado.');
        return response;
      });
  }

  /**
   * Carrega o arquivo .docx para o Docxtemplater
   * @param fileContent Conteúdo do arquivo como ArrayBuffer
   * @returns Instância de Docxtemplater
   */
  private loadDocxFile(fileContent: ArrayBuffer): Docxtemplater {
    try {
      const zip = new PizZip(fileContent);
      return new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
      });
    } catch (error: any) {
      throw new Error('Erro ao carregar o documento: ' + error.message);
    }
  }

  /**
   * Substitui os placeholders no documento .docx
   * @param doc Instância do Docxtemplater com o documento carregado
   * @param placeholders Dados para substituir os placeholders
   * @returns Documento com placeholders substituídos
   */
  private replacePlaceholders(
    doc: Docxtemplater,
    placeholders: Record<string, string>
  ): Docxtemplater {
    try {
      doc.setData(placeholders);
      doc.render();
      return doc;
    } catch (error: any) {
      throw new Error('Erro ao substituir placeholders: ' + error.message);
    }
  }

  /**
   * Salva o arquivo modificado localmente
   * @param doc Documento atualizado
   * @param fileName Nome do arquivo de saída
   */
  private saveDocxFile(doc: Docxtemplater, fileName: string): void {
    const blob = doc.getZip().generate({
      type: 'blob',
      mimeType:
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });
    saveAs(blob, fileName);
  }

  /**
   * Manipulador de erros padrão
   * @param error Objeto de erro
   */
  private handleError(error: any): void {
    console.error('Erro:', error.message);
    alert('Ocorreu um erro ao processar o arquivo: ' + error.message);
  }
}
