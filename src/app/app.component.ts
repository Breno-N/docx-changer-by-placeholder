import { Component, OnInit } from '@angular/core';
import { DocxService } from './docx.service';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css'],
})
export class AppComponent implements OnInit {
  constructor(private docxService: DocxService) {}

  ngOnInit() {
    const fileUrl = '/assets/template.docx';
    const outputFileName = 'template_modificado.docx';

    // Dados dinâmicos para substituir no documento
    const placeholders = {
      sustentabilidade: 'Teste de substituição 1',
      recursosNaturais: 'Teste de substituição 2',
      residuos: 'Teste de substituição 3',
      ambiental: 'Teste de substituição 4',
      social: 'Teste de substituição 5',
      economico: 'Teste de substituição 6',
    };

    this.processDocument(fileUrl, placeholders, outputFileName);
  }

  /**
   * Processa o documento .docx
   * @param fileUrl URL do arquivo a ser processado
   * @param placeholders Dados para substituir os placeholders
   * @param outputFileName Nome do arquivo de saída
   */
  private processDocument(
    fileUrl: string,
    placeholders: Record<string, string>,
    outputFileName: string
  ): void {
    this.docxService.processDocxFile(fileUrl, placeholders, outputFileName);
  }
}
