from random import shuffle

if __name__ == "__main__":
    words = ["Jahe Hangat", "Periksa Tensi", "Vaksin Corona", "Bedah Bedah", "Mata Rabun", "Pernafasan Segar", "Undang-undang Obat"]
    word1 = []
    word2 = []
    for word_pair in words:
        w1,w2 = word_pair.split()
        word1.append(w1)
        word2.append(w2)

    final = []
    for w1 in word1:
        for w2 in word2:
            final.append(f'{w1} {w2}')
    shuffle(final)
    print('\n'.join(final))
    