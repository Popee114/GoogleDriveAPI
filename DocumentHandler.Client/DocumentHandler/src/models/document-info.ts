export class DocumentInfo {
  public id: number;
  public documentId: string;
  public link: string;

  constructor(id: number, documentId: string, link: string) {
    this.id = id;
    this.documentId = documentId;
    this.link = link;
  }
}
