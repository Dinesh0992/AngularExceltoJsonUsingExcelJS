import { async, ComponentFixture, TestBed } from '@angular/core/testing';

import { ExceltojsonComponent } from './exceltojson.component';

describe('ExceltojsonComponent', () => {
  let component: ExceltojsonComponent;
  let fixture: ComponentFixture<ExceltojsonComponent>;

  beforeEach(async(() => {
    TestBed.configureTestingModule({
      declarations: [ ExceltojsonComponent ]
    })
    .compileComponents();
  }));

  beforeEach(() => {
    fixture = TestBed.createComponent(ExceltojsonComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
