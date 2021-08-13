import { Component, OnInit } from '@angular/core';
import { FormControl, FormGroup, Validators } from '@angular/forms';
import { RequestService } from './../services/request.service';
import { DocumentInfo } from './../models/document-info';
import { DomSanitizer, SafeResourceUrl } from '@angular/platform-browser';
import { forkJoin } from 'rxjs';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss'],
  providers: [RequestService]
})
export class AppComponent implements OnInit{
  public title = 'DocumentHandler';

  public firstId: number = 2;

  public secondId: number = 3;

  public isUserExpert: boolean = false;

  public displayModal = false;

  public firstDocLink: SafeResourceUrl = this.sanitize('');

  public secondDocLink: SafeResourceUrl = this.sanitize('');

  public firstDocId!: string;

  public secondDocId!: string;

  public saveChanges: boolean = false;

  public loading: boolean = false;

  private intervalId!: any;

  constructor(private readonly requestService: RequestService, private sanitizer: DomSanitizer) {}

  public ngOnInit(): void {
    this.myForm.valueChanges.subscribe(() => this.onFormChanged());
  }

  public myForm: FormGroup = new FormGroup({
    'firstDocId': new FormControl(this.firstId, Validators.required),
    'secondDocId': new FormControl(this.secondId, Validators.required),
    'isUserExpert': new FormControl(this.isUserExpert)
  });

  public onSubmit(): void {
    this.displayModal = true;
    this.openDocs(this.firstId, this.secondId);
  }

  public closeModal(saveChanges: boolean): void {
    this.saveChanges = saveChanges;
    this.closeDocs(this.firstId, this.firstDocId, this.secondId, this.secondDocId, this.saveChanges);
    clearInterval(this.intervalId);
  }

  public reject(): void {
    this.firstDocLink = this.sanitize('assets/loading-page.html');
    this.secondDocLink = this.sanitize('assets/loading-page.html');
    forkJoin(
      {
        firstDocument: this.requestService.rejectDocument(this.firstId, this.firstDocId),
        secondDocument: this.requestService.rejectDocument(this.secondId, this.secondDocId)
      }
    ).subscribe(
      (data) => {
        clearInterval(this.intervalId);
        this.displayModal = false;
        this.loading = false;
        console.log(data);
      },
      () => {
        clearInterval(this.intervalId);
        this.displayModal = false;
        this.loading = false;
      }
    );
  }

  public register(): void {
    alert('Регистрация документов выполнена');
    this.displayModal = false;
  }

  private openDocs(firstId: number, secondId: number): void {
    this.loading = true;
    this.firstDocLink = this.sanitize('assets/loading-page.html');
    this.secondDocLink = this.sanitize('assets/loading-page.html');
    this.openingDocsRequest(firstId, secondId);
    this.intervalId = setInterval(() => {
      console.clear();
    }, 1000);
  }

  private closeDocs(firstId: number, firstDocId: string, secondId: number, secondDocId: string, saveChanges: boolean): void {
    this.firstDocLink = this.sanitize('assets/loading-page.html');
    this.secondDocLink = this.sanitize('assets/loading-page.html');
    if (saveChanges) {
      this.closingDocsRequest(firstId, firstDocId, secondId, secondDocId);
    } else {
      this.displayModal = false;
    }
  }

  private openingDocsRequest(firstId: number, secondId: number) {
    forkJoin(
      {
        firstDocument: this.requestService.openDocument<DocumentInfo>(firstId),
        secondDocument: this.requestService.openDocument<DocumentInfo>(secondId)
      }
    ).subscribe(
      (data) => {
        this.firstDocLink = this.sanitize(data.firstDocument.link);
        this.firstDocId = data.firstDocument.documentId;
        this.secondDocLink = this.sanitize(data.secondDocument.link);
        this.secondDocId = data.secondDocument.documentId;
        this.loading = false;
      },
      () => {
        this.firstDocLink = this.sanitize('assets/error-page.html');
        this.secondDocLink = this.sanitize('assets/error-page.html');
        this.loading = false;
      }
    )
  }

  private closingDocsRequest(firstId: number, firstDocId: string, secondId: number, secondDocId: string): void {
    forkJoin(
      {
        firstDocument: this.requestService.closeDocument(firstId, firstDocId),
        secondDocument: this.requestService.closeDocument(secondId, secondDocId)
      }
    ).subscribe(
      () => {
        this.displayModal = false;
        this.loading = false;
      },
      () => {
        this.displayModal = false;
        this.loading = false;
      }
    );
  }

  private onFormChanged(): void {
    this.firstId = this.myForm.get('firstDocId')?.value;
    this.secondId = this.myForm.get('secondDocId')?.value;
    this.isUserExpert = this.myForm.get('isUserExpert')?.value;
  }

  private sanitize(url: string): SafeResourceUrl {
    return this.sanitizer.bypassSecurityTrustResourceUrl(url);
  }
}
