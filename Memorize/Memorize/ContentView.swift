//
//  ContentView.swift
//  Memorize
//
//  Created by Valerio  Biolchi on 02/11/2020.
//

import SwiftUI
// variables, functions, behaviors containers
struct ContentView: View { //behavior specification
    var body: some View {
        HStack {
            ForEach(0..<4) { index in
                CardView()
            }
        }
            .padding()
            .foregroundColor(.orange)
            .font(Font.largeTitle)
    }
}

struct CardView: View {
    var isFaceUp: Bool = true
    var body: some View {
        ZStack {
            if isFaceUp {
                RoundedRectangle(cornerRadius: 10.0).fill(Color.blue)
                RoundedRectangle(cornerRadius: 10.0).stroke(lineWidth: 3)
                Text("🏀")
            } else {
                RoundedRectangle(cornerRadius: 10.0).fill()
            }
        }
    }
}














struct ContentView_Previews: PreviewProvider {
    static var previews: some View {
        ContentView()
    }
}
