import { Injectable } from '@angular/core';
import { HttpClient, HttpParams } from '@angular/common/http';
import { Observable } from 'rxjs';
import { environment } from 'src/environments/environment';
import { settingsConfig } from './../models/settings.config';

@Injectable()
export class RequestService{

  constructor(private httpClient: HttpClient){}

  public openDocument<T>(id: number): Observable<T> {
    const params = new HttpParams().set('fileId', id);
    return this.httpClient.get<T>(this.getHost() + settingsConfig.api.open, { params });
  }

  public closeDocument<T>(id: number, docId: string) {
    const body = {
      fileId: id,
      documentId: docId
    }
    return this.httpClient.post<T>(this.getHost() + settingsConfig.api.close, body);
  }

  public rejectDocument<T>(id: number, docId: string) {
    const body = {
      fileId: id,
      documentId: docId
    }
    return this.httpClient.post<T>(this.getHost() + settingsConfig.api.reject, body);
  }

  private getHost(): string {
    return environment.host;
  }
}
