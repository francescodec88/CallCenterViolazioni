import { Component } from '@angular/core';
import { CommonModule } from '@angular/common';
import { MatCardModule } from '@angular/material/card';
import { BrowserAnimationsModule } from '@angular/platform-browser/animations';
import { trigger, transition, style, animate } from '@angular/animations';

interface CardItem {
  title: string;
  content: string;
}

@Component({
  selector: 'app-dashboard',
  standalone: true,
  imports: [CommonModule, MatCardModule, BrowserAnimationsModule],
  templateUrl: './dashboard.html',
  styleUrls: ['./dashboard.scss'],
  animations: [
    trigger('fadeSlide', [
      transition(':enter', [
        style({ opacity: 0, transform: 'translateY(20px)' }),
        animate('600ms ease-out', style({ opacity: 1, transform: 'translateY(0)' }))
      ])
    ])
  ]
})
export class DashboardComponent {
  cards: CardItem[] = [
    { title: 'Statistiche', content: 'Overview delle metriche' },
    { title: 'Segnalazioni', content: 'Lista segnalazioni recenti' },
    { title: 'Utenti', content: 'Gestione utenti' },
    { title: 'Impostazioni', content: 'Configurazioni avanzate' },
    { title: 'Report', content: 'Generazione report' },
    { title: 'Assistenza', content: 'Contatti e supporto' }
  ];
}
